del bin\*.xlsm
copy template\DBTBL_postgres.xlsm bin\DBTBL.xlsm
cscript vbac.wsf combine

mv bin\DBTBL.xlsm bin\DBTBL_postgres.xlsm

copy template\DBTBL_mysql.xlsm bin\DBTBL.xlsm
cscript vbac.wsf combine

mv bin\DBTBL.xlsm bin\DBTBL_mysql.xlsm
