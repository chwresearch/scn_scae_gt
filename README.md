# scn_scae_gt

Este repositorio contiene los archivos de procesamiento y análisis de la Cuenta 
de Energía del Sistema de Contabilidad Ambiental y Económica de Guatemala. 

El proyecto completo se ha desarrollado utilizando R como herramienta principal. 
Por conveniencia se ha utilizado la utilidad de "Proyectos" de `RStudio`, la 
cual permite que los archivos de procesamiento hagan referencia a otros archivos 
de manera relativa, eliminando la necesidad de establecer el directorio de 
trabajo a través de `getwd()`. Por eso, se recomienda utilizar `RStudio` y abrir 
este proyecto a través del archivo `scn_scae_gt.Rproj`.


```
│   .gitignore
│   README.html
│   README.md
│   scn_scae_gt.Rproj    <---- ABRIR ESTE ARCHIVO EN RSTUDIO.
│
├───MemoriaDeAnalisis
│       MemoriaDeAnalisis.docx
│       MemoriaDeAnalisis.html
│       MemoriaDeAnalisis.pdf
│       MemoriaDeAnalisis.qmd
│ 
│
└───scripts
    │   CuentaDeEnergia.R
    │   Procesamiento_COUs.R
    │   VIEW_oferta_utilizacion.sql
    │
    └───datos
            .Rhistory
            CLASIFICACIONES.xlsx
            GTM_COUS_DESAGREGADOS_2013_2020.xlsx
            scn.db
            SCN_BD.xlsx

```
