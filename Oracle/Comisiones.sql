SELECT * FROM USRFINDATA.FT_COMISION_MIG@BSIBPDW0@UFINRIE FC
INNER JOIN USRFINDATA.ODS_PARAMETRO@BSIBPDW0@UFINRIE OP ON OP.COD_CLASE='BCINSER' AND
                                       OP.COD_TIPPARAM='CTAR'    AND
                                       RPAD(FC.COD_CUENTA,15,'0')=OP.COD_PARAMINI
JOIN USRFINDATA.ODS_OFICINA@BSIBPDW0@UFINRIE B ON FC.COD_OFICINA=B.COD_OFICINA AND B.FECPRO='30/06/2022' AND B.COD_BANCA IN ('2685','2661')
LEFT JOIN (           
           SELECT * FROM USRFINDATA.stg_pe01@BSIBPDW0@UFINRIE WHERE fecpro=(SELECT MAX(fecpro) FROM USRFINDATA.stg_pe01@BSIBPDW0@UFINRIE ) 
           ) C ON FC.COD_CLIENTE=C.CENTRAL                                       
-- MODIFICAR FEC_PROCESO SEG�N EL PERIODO REQUERIDO
WHERE FC.FEC_PROCESO > '01/01/2019'  AND DES_CONCEPTO = 'Fianzas'
AND COD_EPIGRAFE LIKE '%01'


SELECT * FROM USRFINDATA.ODS_OFICINA@BSIBPDW0@UFINRIE
WHERE COD_BANCA IN ('2685','2661')