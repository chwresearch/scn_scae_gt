library("duckdb")

if (file.exists("scn_gt.duckdb")) {
  file.remove("scn_gt.duckdb")
}

con <- dbConnect(duckdb(), dbdir = "salidas/scn_gt.duckdb", read_only = FALSE)
dbExecute(con, "INSTALL sqlite;")
dbExecute(con, "LOAD sqlite;")
dbExecute(con, "CREATE TABLE scn AS (SELECT * FROM sqlite_scan('salidas/scn.db', 'scn'));")
dbGetQuery(con, "SHOW TABLES;")
dbDisconnect(con)