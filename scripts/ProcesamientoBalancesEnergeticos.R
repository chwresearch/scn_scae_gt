# Llamar librerías
library(readxl)
library(openxlsx)
library(reshape2)
library(stringr)
library(plyr)

# Limpiar el área de trabajo
rm(list = ls())


# Lógica recursiva
# ================

# Limpiar el área de trabajo
rm(list = ls())


archivo <- "datos/GTM_Balances_Energeticos_2013-2020.xlsx"
hojas <- excel_sheets(archivo)

lista <- c("inicio")

for (i in 1:length(hojas)){
  # Extraemos el año y la unidad de medida
  info <- read_excel(
    archivo,
    range = paste("'" , hojas[i], "'!A4:B6", sep = "") ,
    col_names = FALSE,
    col_types = "text",
  )
  
  # Extraemos el texto de la cadena de caracteres
  anio <- as.numeric((str_extract(info[1,1 ], "\\d{4}")))
  unidad_olade <- toString(info[3, 2]) # unidad de medida
  unidad <- "Terajoules"
  cuadro <- "Balance Energético"
  
  # Balance energetico total
  # ================
  
  balance <- as.matrix(read_excel(
    archivo,
    range = paste("'" , hojas[i], "'!B7:X33", sep = ""),
    col_names = FALSE,
    col_types = "numeric"
  ))
  
  balance[is.na(balance)] <- 0.0 # cambiamos "sin dato" a cero por conveniencia 
  balance[c(7:14),][balance[c(7:14),]>0]  <- 0 # cambiamos los positivos que no suman a la oferta
  balance[c(4,7:14),] <- balance[c(4,7:14),] * (-1) # cambiamos los negativos
  rownames(balance) <- c(sprintf("bf%03d", seq(1, dim(balance)[1])))
  colnames(balance) <- c(sprintf("bc%03d", seq(1, dim(balance)[2])))
  
  # Filas a eliminar con subtotales y totales
  
  # 6	  OFERTA TOTAL
  # 15	TOTAL TRANSFORMACIÓN
  # 25	CONSUMO ENERGÉTICO
  # 27  CONSUMO FINAL
  
  # Columnas a eliminar con subtotales y totales
  
  # 10	TOTAL PRIMARIAS
  # 22	TOTAL SECUNDARIAS
  # 23	TOTAL
  
  # Y se eliminan los datos calculados (duplicados)
  balance <- balance[-c(6,15,25,27),-c(10,22,23)]
  
  # Desdoblamos
  balanceBD <- cbind(anio,cuadro, melt(balance), unidad)
  
  # Renombramos filas y columnas
  colnames(balanceBD) <- c("anio", "cuadro", "correlativo_filas", 
                         "correlativo_columnas", "valor", "unidad")
  
  # Y guardamos el objeto completo
  assign(paste("BAL_",anio, sep = ""), balanceBD)
  
  # Modificamos nuestra lista de objetos
  lista <- c(lista, paste("BAL_",anio , sep = ""))
}  

# Actualizamos nuestra lista de objetos creados
lista <- lapply(lista[-1], as.name)

# Unimos los objetos de todos los años y precios
BAL <- do.call(rbind.data.frame, lista)

# Y borramos los objetos individuales
do.call(rm,lista)

  
# Le damos significado a las filas y columnas
  
clasificacionColumnas <- read_xlsx(
    "datos/CLASIFICACIONES_BALANCES.xlsx",
    sheet = "columnas",
    col_names = TRUE
  )
  
clasificacionFilas <- read_xlsx(
    "datos/CLASIFICACIONES_BALANCES.xlsx",
    sheet = "filas",
    col_names = TRUE)

clasificacionFilasXcolumnas <- read_xlsx(
  "datos/CLASIFICACIONES_BALANCES.xlsx",
  sheet = "filasXcolumnas",
  col_names = TRUE)

BAL <- join(BAL,clasificacionFilas, by = "correlativo_filas")
BAL <- join(BAL,clasificacionColumnas,by = "correlativo_columnas")

# Creamos la concatenacion de filas columnas para el ambiente
BAL$id_filasXcolumnas <- paste(BAL$id_oferta_utilizacion, BAL$id_ntge, BAL$id_energetico, sep = "")

# Y le damos sentido a la concatenación
BAL <- join(BAL,clasificacionFilasXcolumnas, by = "id_filasXcolumnas")

gc()

con <- dbConnect(RSQLite::SQLite(), "datos/scn.db")
dbCreateTable(con, "balances_energeticos", BAL)
dbAppendTable(con, "balances_energeticos", BAL)
dbDisconnect(con)

# Y lo exportamos a Excel
write.xlsx(
  BAL,
  "datos/BALGT_BD.xlsx",
  sheetName= "BALGT_BD",
  rowNames=FALSE,
  colnames=FALSE,
  overwrite = TRUE,
  asTable = FALSE
)

# El formato CSV se exporta muy grande, pero se comprime muy bien a 3mb
write.csv(
  BAL,
  "datos/balgt.csv",
  row.names = FALSE,
  fileEncoding = "UTF-8"
)
