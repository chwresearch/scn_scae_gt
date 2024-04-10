library("duckdb")
con <- dbConnect(duckdb(), dbdir = "salidas/scn_gt.duckdb", read_only = FALSE)
#dbExecute(con, "INSTALL sqlite;")
#dbExecute(con, "LOAD sqlite;")
dbGetQuery(con, "ATTACH 'salidas/scn.db' AS test (TYPE sqlite);")
dbGetQuery(con, "SHOW TABLES;")

dbDisconnect(con)