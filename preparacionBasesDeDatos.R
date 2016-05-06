require(data.table)
require(openxlsx)
rm(list=ls())
# preparacion bases de datos

# Diagnósticos
# dx=data.table(read.xlsx("/Users/aldopareja/Google Drive/CCB/BD/ReporteDiagnosticos_25-04-2016-112831.xlsx"))
# save(dx,file="backUps/diagnosticos.RData")
load("backUps/diagnosticos.RData")
dx=dx[,list(ID.DX,Fecha.de.Creación.DX,Nombre.DX,Porcentaje.DX,Estado.DX,Ciudad.Usuario,Ciudad.Empresa,
            Tipo.Documento.Usuario,Documento.Usuario,Tipo.Documento.Empresa,Razón.Social
            ,Genero.Usuario,Edad.Usuario,Nivel.Estudio.Usuario,CIIU.DX,Nombre.CIIU.DX,
            Sector.DX,Usuario.Asignado,Area.Asesor.Linea.Asesor.Asignado)]
write.xlsx(dx,"BD/diagnosticos.xlsx",asTable = T)

# Rutas
# rutas=data.table(read.xlsx("/Users/aldopareja/Google Drive/CCB/BD/ReporteRutaServicios_03-05-2016-050522.xlsx"))
# save(rutas,file="backUps/rutas.RData")
load("backUps/rutas.RData")
write.xlsx(rutas,"BD/rutas.xlsx",asTable = T)

# Infoservicios

rutas=data.table(read.xlsx("/Users/aldopareja/Google Drive/CCB/BD/ReporteRutaServicios_03-05-2016-050522.xlsx"))
save(rutas,file="backUps/rutas.RData")
