require(data.table)
require(openxlsx)
rm(list=ls())
#I
#filtros
#base Dx
db=data.table(read.xlsx("BD/diagnosticos.xlsx"))
#Diagnósticos diligenciados
db=db[Estado.DX=="Diligenciado"]
#ignorar formalizacion y nueva línea de negocio
db=db[!grepl("Forma|Nueva",Nombre.DX)]
#solo Dx en la jurisdiccion
jrcion=read.xlsx("BD/jurisdiccion.xlsx",colNames = F)$X1
#quitar las comas, cambiar el encoding, quitar apóstrofes, quitar espacios al final
jrcion=toupper(gsub(",|'|[[:space:]]$","",iconv(jrcion,to="ASCII//TRANSLIT")))
jrcion[jrcion=="ZIPAQUIRA."]="ZIPAQUIRA"
jrcion[jrcion=="VENECIA"]="VENECIA - OSPINA PEREZ"
#juntar las variables ciudad cuando la de la empresa no está
db[is.na(Ciudad.Empresa),Ciudad.Empresa:=Ciudad.Usuario]
#Solo Dx en jurisdiccion
db=db[Ciudad.Empresa%in%jrcion]
#Dx únicos por cliente
#setear la fecha como tipo fecha
db[,Fecha.de.Creación.DX:=as.Date(Fecha.de.Creación.DX)]
#se deja el Dx más reciente por cada número de cédula
db=db[,.SD[which.max(Fecha.de.Creación.DX)]
      ,by=Documento.Usuario]
#creacion variable sector
db[,area:="multisectorial"]
db[Sector.DX%in%c("AGROINDUSTRIAL","AGRÍCOLA"),area:="agroindustrial"]
db[Sector.DX=="TEXTIL Y CONFECCIÓN",area:="confecciones"]
db[Sector.DX=="INDUSTRIAS CREATIVAS Y CULTURALES",area:="ICC"]

#leo base de rutas
rutas=data.table(read.xlsx("BD/rutas.xlsx"))
#cruzo con la base de Dx
db=merge(db,rutas[,list(ID_DX,ESTADO_ASISTENCIA_ACTIVIDAD)]
         ,by.x="ID.DX",by.y="ID_DX",all.x=F,all.y = F)

#calculo el porcentaje de avance en ruta
db=unique(db[,porcAvanceRuta:=round(100*sum(ESTADO_ASISTENCIA_ACTIVIDAD=="ASISTIO"
                                            ,na.rm = T)/.N)
             ,by=ID.DX])
#a hace un histograma del porcentaje de avance en ruta
hist(db[,porcAvanceRuta],100)

#b
#agrego por area y nombre de Dx y hayo el promedio
table=db[,list(`porcentaje avance promedio`=round(mean(porcAvanceRuta),2))
         ,by='area,Nombre.DX']
#hago un cast para que me quede mejor visualizado
table=dcast(table,area~Nombre.DX)

#escribo una tabla en PDF
require(gridExtra)
pdf("Respuestas/promediosAvanceEnRuta.pdf",height = 2)
grid.table(table,rows=NULL)
dev.off()

#parte II
rm(list=ls())
#leo cada base desde la línea 10 y las pego
prog<-data.table(read.xlsx("BD/programaciones-asistencias/Programación Febrero - Marzo 2015.xlsx"
                           ,sheet = "Formato Consolidado",startRow = 10))
#abril-mayo
temp<-data.table(read.xlsx("BD/programaciones-asistencias/Programación Abril - Mayo 2015.xlsx"
                           ,sheet = "Formato Consolidado",startRow = 10))
prog<-rbind(prog,temp)
#junio-julio
temp<-data.table(read.xlsx("BD/programaciones-asistencias/Programación Junio - Julio 2015.xlsx"
                           ,sheet = "Formato Consolidado",startRow = 10))
prog<-rbind(prog,temp)
#agosto
temp<-data.table(read.xlsx("BD/programaciones-asistencias/Programación Agosto 2015.xlsx"
                           ,sheet = "Formato Consolidado",startRow = 10))
prog<-rbind(prog,temp)
#septiembre
temp<-data.table(read.xlsx("BD/programaciones-asistencias/Programación Septiembre 2015.xlsx"
                           ,sheet = "Formato Consolidado",startRow = 10))
prog<-rbind(prog,temp)
#oct-nov
temp<-data.table(read.xlsx("BD/programaciones-asistencias/Programación Octubre - Noviembre 2015.xlsx"
                           ,sheet = "Formato Consolidado",startRow = 10))
temp<-rbind(prog,temp)
rm(prog)

#arreglo los códigos y saco los servicios que no tengan código
#quito los espacios
temp[,CODIGO:=gsub("[[:space:]]","",CODIGO)]
#quito estos caracteres raros
temp[,CODIGO:=gsub("\r\n","",CODIGO)]
temp=temp[!(CODIGO=="NA"|is.na(CODIGO))]
temp=temp[is.na(`CANCELADAS./.REPROGRAMADAS`)|`CANCELADAS./.REPROGRAMADAS`==" "|!grepl("CANCEL",DIA)]
#analisis de horarios
#el análisis se hace solo para talleres, cápsulas y asesorías grupales
#solo para las sedes de chapi, kenne y salitre
#solo actividades abiertas y presenciales
horario=temp[grepl("Taller|conocimiento|grupal",TIPO.DE.SERVICIO)
             &grepl("kenne|chapine|salitre",SEDE,ignore.case = T)
             &ASISTENTES>0
             &grepl("ABIER",`ABIERTA./.CERRADA`,ignore.case = T)
             &grepl("presencial",MODALIDAD.DE.ENTREGA,ignore.case = T)]
rm(temp)
#a
#Creo variable (mañana, tarde y noche)
#Modifico el dia para que sea homogéneo
#modifico abierta/cerrada para que quede homogéneo
horario=horario[,c("sesion","DIA","ABIERTA./.CERRADA","SEDE"):={
                      sesion=sapply(strsplit(HORA,""), function(x) which(x=='.')[2])
                      sesion=substr(HORA,1,sesion-1)
                      temp=sapply(strsplit(HORA,""),function(x) which(x==':')[1])
                      temp=as.integer(substr(sesion,1,temp-1))
                      sesion=ifelse(grepl('p',sesion),temp+12,temp)
                      sesion=ifelse(sesion<12,'mañana',ifelse(sesion>=12&sesion<17,'tarde','noche'))
                      DIA=gsub("'","",iconv(tolower(DIA),to = "ASCII//TRANSLIT"))
                      `ABIERTA./.CERRADA`=ifelse(grepl("ABIER",`ABIERTA./.CERRADA`),'abierta','cerrada')
                      sede=tolower(gsub("[[:space:]]$","",gsub("[[:space:]]$","",SEDE)))
                      list(sesion,DIA,`ABIERTA./.CERRADA`,sede)
                      }]
#agrego por día y hora
temp=horario[,list(programaciones=.N,
                  asistencias=sum(.SD[,ASISTENTES],na.rm = T),
                  promAsist=round(sum(.SD[,ASISTENTES],na.rm = T)/.N,2))
                ,by=c("sesion","DIA")][order(-promAsist)][programaciones>5]

#escribo una table bonita y ordenada
require(gridExtra)
pdf("Respuestas/horariosMasFavorables.pdf",width = 5.9,height = 4.2)
grid.table(temp,rows=NULL)
dev.off()

#b
#hago la regresión lineal
#guardo la variable explicativa
x=horario[TOTAL.INSCRITOS.POR.PAGINA.WEB>0,
          TOTAL.INSCRITOS.POR.PAGINA.WEB]
# guardo la variable explicada
y=horario[TOTAL.INSCRITOS.POR.PAGINA.WEB>0,
          ASISTENTES]
# ajusto un modelo lineal
mylm<-lm(y~0+x)
# hago una predicción al 70% de confianza del modelo
temp=predict(mylm,interval="prediction",level = 0.7)
#grafico
plot(horario[TOTAL.INSCRITOS.POR.PAGINA.WEB>0,
             list(TOTAL.INSCRITOS.POR.PAGINA.WEB,ASISTENTES)])
#pongo la línea ajustada promedio
lines(x,temp[,1],lwd=2,col='blue')
#agrego el intervalo superior
lines(x,temp[,3],lwd=2,col='red')
dev.off()

summary(mylm)
intSuperior=unique(round(temp[,3]-temp[,1],1))

#calculo los cupos de acuerdo con el intervalo
ceiling((30-intSuperior)/as.numeric(mylm["coefficients"]))
