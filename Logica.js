
document.addEventListener("DOMContentLoaded", () => {
  let categoriaActiva = null;
  let filaContador = 0;
  let lastAddTime = 0; // para evitar clicks rápidos

  // ----------------------------- CATALOGO -----------------------------
  // (usa aquí tu catálogo completo tal como lo tenías)
  const catalogo = {
    insumos: [
      { clave: "S/C", descripcion: "BENZOCAÍNA 20% GEL, FRASCO 30 g", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "MEPIVACAÍNA 3% SIN VASOCONSTRICTOR. CAJA CON 50 CARTUCHOS DENTALES DE 1.8 ML CADA UNO.", stock: "", minimo: "", caducidad: "" },
      { clave: "010.000.0267.00", descripcion: "LIDOCAÍNA, EPINEFRINA. SOLUCIÓN INYECTABLE AL 2%. CAJA CON 50 CARTUCHOS DENTALES CON 1.8 ML", stock: "", minimo: "", caducidad: "" },
      { clave: "070.817.0550", descripcion: "SOLUCIÓN PARA REVELADO PARA SISTEMA MANUAL DE REVELADOR CONCENTRADO. ENVASE 828 ML. (LA MARCA DE ESTE PRODUCTO DEBE SER LA MISMA DE LA CLAVE 070.426.0363)1", stock: "", minimo: "", caducidad: "" },
      { clave: "070.426.0363", descripcion: "SOLUCIÓN PARA FIJADO PARA SISTEMA MANUAL DE FIJADOR CONCENTRADO. ENVASE 828 ML. (LA MARCA DE ESTE PRODUCTO DEBE SER LA MISMA DE LA CLAVE 070.817.0550)1", stock: "", minimo: "", caducidad: "" },
      { clave: "070.707.0496", descripcion: "PELÍCULA RADIOGRÁFICA, DENTAL ADULTO, MEDIDAS: EN UN RANGO DE 3 A 3.5 CM POR 4 A 4.5 CM. CAJA CON 150 PELÍCULAS1", stock: "", minimo: "", caducidad: "" },
      { clave: "070.707.0587", descripcion: "PELÍCULA RADIOGRÁFICA INFANTIL SENCILLA PERIAPICAL DE 2.2 * 3.5 CM. CAJA CON 100 PELÍCULAS1", stock: "", minimo: "", caducidad: "" }
    ],
    material: [
      { clave: "060.016.0204", descripcion: "ACEITES . LUBRICANTE PARA PIEZA DE MANO DE BAJA VELOCIDAD. ENVASE CON APLICADOR CON 120 ML.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.016.0253", descripcion: "ACEITES. ACEITE LUBRICANTE PARA TURBINA DE PIEZA DE MANO DE ALTA VELOCIDAD. APLICADOR EN FORMA DE JERINGA. ENVASE CON 2 ML.", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "ÁCIDO FOSFÓRICO AL 34% O 37% PARA GRABADO DE ESMALTE O DENTINA.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.031.0072", descripcion: "ADHESIVO. ADHESIVO DENTAL PARA RESINAS DIRECTAS AUTOPOLIMERIZABLE O FOTOPOLIMERIZABLE. FRASCO DE 5 ML.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.040.8058", descripcion: "AGUJAS DENTAL .TIPO CARPULE. DESECHABLE. LONGITUD. 25-42 MM CALIBRE. 27 G TAMAÑO. LARGA.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.040.8041", descripcion: "AGUJAS DENTALES. TIPO CARPULE. DESECHABLE. LONGITUD. 20-25 MM CALIBRE: 30 G TAMAÑO: CORTA.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.797.0019", descripcion: "ALGODONES. PARA USO DENTAL. MEDIDA: 3.8 X 0.8 CM. ENVASE CON 500 ROLLOS.", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "ANTISÉPTICO BUCAL, SOLUCIÓN ELECTROLIZADA CON PH NEUTRO AL 0.0015% (15PPM) DE CL ACTIVO. CADA 100 ML. CONTIENE: IONES ACTIVOS: 7.5 - 9.5 MG. RANGO DE PH DE 6.4 A 7.5. ORP DE 650 A 900 MV", stock: "", minimo: "", caducidad: "" },
      { clave: "060.066.0401", descripcion: "ANTISÉPTICOS. EUGENOL.", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "APLICADORES MULTIPROPÓSITO (MICROBRUSH) CON PUNTAS DE FIBRA NO ABSORBENTE Y CUELLO FLEXIBLE. TAMAÑO: FINO, SUPERFINO O REGULAR.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.111.0208", descripcion: "BARNICES. DE COPAL. PARA REVESTIMIENTO DE CAVIDADES.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.066.1078", descripcion: "BARNIZ DE FLUORURO DE SODIO AL 5% EN UNA CONCENTRACIÓN DE 22600 PPM AUTOPOLIMERIZABLE, EN UN VEHÍCULO DE RESINA MODIFICADO.", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "CAMPO DESECHABLE. CUENTA CON DOS CAPAS DE PROTECCIÓN: 1 CAPA DE PAPEL Y UNA DE POLIETILENO. MED. 45 CM X 35 CM, DIFERENTES COLORES", stock: "", minimo: "", caducidad: "" },
      { clave: "060.182.0178", descripcion: "CEMENTO DE IONOMERO, DE VIDRIO RESTAURATIVO II. COLOR NO 21. POLVO 15 G. SILICATO DE ALUMINIO 95% - 97%. ACIDO POLIACRILICO 3% - 5%. LIQUIDO: 10 G, 8 ML. ACIDO POLIACRILICO 75%. ACIDO TARTÁRICO 10% - 15%. BARNIZ COMPATIBLE LIQUIDO 10 G.", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "CEMENTO DENTAL. PROTECTOR PULPAR PARA BASE. HIDRÓXIDO DE CALCIO COMPUESTO FOTOPOLIMERIZABLE.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.182.0160", descripcion: `CEMENTO: IONÓMERO DE VIDRIO I.
    PARA CEMENTACIONES DEFINITIVAS.
    POLVO 35 G.
    SILICATO DE ALUMINIO 95% -97%.
    ÁCIDO POLIACRÍLICO 3% - 5%.
    LÍQUIDO 25 G, 20 ML.
    ÁCIDO POLIACRÍLICO 75%.
    ÁCIDO POLIBÁSICO 10-15%.
    JUEGO.`, stock: "", minimo: "", caducidad: "" },
      { clave: "060.182.1366", descripcion: "CEMENTOS DENTALES. PARA USO QUIRÚRGICO. DE OXIDO DE ZINC CON ENDURECEDOR (POLVO) 65 G Y EUGENOL (LIQUIDO) 30 ML. CON GOTERO DE PLÁSTICO.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.182.1150", descripcion: `CEMENTOS PROTECTOR PULPAR PARA SELLAR CAVIDADES DENTALES. DE HIDRÓXIDO DE CALCIO, COMPUESTO Y APLICADOR AUTOPOLIMERIZABLE, DOS PASTAS SEMILÍQUIDAS, BASE 13 DESECHABLE G Y CATALIZADOR 11 G CON BLOQUE DE PAPEL PARA MEZCLAR.`, stock: "", minimo: "", caducidad: "" },
      { clave: "060.182.0236", descripcion: `CEMENTOS. DENTAL PARA USO QUIRÚRGICO. POLVO ÓXIDO DE ZINC. POLVO ROSA. TALCO. LÍQUIDO: EUGENOL. ALCOHOL ISOPROPÍLICO AL 10%. RESINA DE PINO. ACEITE DE PINO. ACEITE DE CLAVO. ACEITE DE CACAHUATE. ESTUCHE.`, stock: "", minimo: "", caducidad: "" },
      { clave: "060.182.1176", descripcion: "CEMENTOS. DENTALES. DE OXIFOSFATO DE ZINC. POLVO Y LÍQUIDO. CAJA CON 32 G DE POLVO Y 15 ML DE SOLVENTE. ESTUCHE.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.182.1275", descripcion: "CEMENTOS. DENTALES. PARA RESTAURACIÓN INTERMEDIA. DE ÓXIDO DE ZINC (POLVO) 38 G Y EUGENOL (LÍQUIDO) 14 ML. CON GOTERO DE PLÁSTICO. JUEGO.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.182.0194", descripcion: "CEMENTOS. IONÓMERO DE VIDRIO CON ALEACIÓN DE LIMADURA DE PLATA. POLVO 15 G. SILICATO DE ALUMINIO 100%. LÍQUIDO 10 G 8 ML.  ÁCIDO POLIACRÍLICO 45%. POLVO DE LIMADURA DE PLATA 17 G. PLATA 56% ESTAÑO 29% COBRE 15%.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.182.1424", descripcion: "CEMENTOS. IONÓMERO DE VIDRIO I. PARA CEMENTACIONES DEFINITIVAS. POLVO 35 G. VIDRIO DE SILICATO DE ALUMINIO TRATADO CON SILANO 98%. POLÍMERO DE CELULOSA < 1.5%. DIÓXIDO DE TITANIO < 1%. CATALIZADOR PARA POLIMERIZACIÓN <1%. LÍQUIDO 22.5 ML. COPO LIMERO DE ÁCIDO ACRÍLICO Y ÁCIDO ITACÓNICO 35%. AGUA 35%. 2-HIDROXIETIL-METACRILATO 30%. ÁCIDO DI CARBOXÍLICO 1%.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.182.0186", descripcion: "CEMENTOS. IONÓMERO DE VIDRIO RESTAURATIVO II.COLOR NO. 22.POLVO 15 G.SILICATO DE ALUMINIO 95% -97%.ÁCIDO POLIACRÍLICO 3% -5%.LÍQUIDO 10 G 8 ML.ÁCIDO POLIACRÍLICO 75%.ÁCIDO TARTÁRICO 10% - 15%.BARNIZ COMPATIBLE LÍQUIDO 10 G. JUEGO.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.182.1440", descripcion: "CEMENTOS. IONÓMERO DE VIDRIO RESTAURATIVO TIPO II. PARA TRATAMIENTO RESTAURATIVO ATRAUMÁTICO (TRA). PARA RESTAURACIONES INTERMEDIAS. PARA BASES. PARA ODONTOLOGÍA MÍNIMAMENTE INVASIVA (OMI). TONO A3 . POLVO GRANULADO RADIOPACO: 12.5 G. VIDRIO DE FLUOROSILICATO DE CALCIO LANTANO. ALUMINIO RECUBIERTO 90% . ÁCIDO POLIACRÍLICO. 10% ÁCIDO BENZÓICO <0.1% PIGMENTOS <0.1% LÍQUIDO DE 8.5 ML (10GR). AGUA  55%-65%  COPOLÍMERO  DE  ÁCIDO ACRÍLICO  Y ÁCIDO MALÉICO. 25-35% ÁCIDO TARTÁRICO. 9.1% ÁCIDO BENZÓICO. 0.1% LOSETA DE PAPEL ENCERADO CUCHARILLA DISPENSADORA GUÍA DE APLICACIÓN E INSTRUCTIVO. ESTUCHE.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.182.0228", descripcion: "CEMENTOS. IONÓMERO DE VIDRIO. PARA TRATAMIENTO RESTAURATIVO ATRAUMÁTICO. POLVO: 10 G. SILICATO DE ALUMINIO 89 -95%. ÁCIDO POLIACRÍLICO 0.-10%. LÍQUIDO 6 G 4.8 ML AGUA DESTILADA.ÁCIDO POLIACRÍLICO 40 -50%.BARNIZ 5 G.CLORURO DE POLIVINIL 10 -20%.ACETATO ETÍLICO 75 -85%. ESTUCHE.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.189.0254", descripcion: "CEPILLO PARA PULIDO DE AMALGAMAS Y PROFILAXIS. DE CERDAS BLANCAS EN FORMA DE COPA. PARA CONTRA-ÁNGULO.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.189.0106", descripcion: "CEPILLOS. DENTAL INFANTIL, CON MANGO DE PLÁSTICO Y CERDAS RECTAS DE NYLON 6.12, 100 % VIRGEN O POLIÉSTER P.B.T. 100 % VIRGEN, DE PUNTAS REDONDEADAS EN 3 HILERAS, CABEZA CORTA, CONSISTENCIA MEDIANA.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.189.0015", descripcion: "CEPILLOS. DENTAL, PARA ADULTO, CON MANGO DE PLÁSTICO Y CERDAS RECTAS DE NYLON 6.12, 100 % VIRGEN O POLIÉSTER P.B.T. 100 % VIRGEN, DE PUNTAS REDONDEADAS EN 4 HILERAS, CABEZA CORTA, CONSISTENCIA MEDIANA.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.189.0023", descripcion: "CEPILLOS. PARA PULIDO DE AMALGAMAS Y PROFILAXIS. DE CERDAS BLANCAS EN FORMA DE COPA. PARA PIEZA DE MANO. PIEZA", stock: "", minimo: "", caducidad: "" },
      { clave: "060.189.0031", descripcion: "CEPILLOS. PARA PULIDO DE AMALGAMAS Y PROFILAXIS. DE CERDAS NEGRAS EN FORMA DE BROCHA. PARA CONTRA-ÁNGULO. PIEZA.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.189.0205", descripcion: "CEPILLOS. PARA PULIDO DE AMALGAMAS Y PROFILAXIS. DE CERDAS NEGRAS EN FORMA DE BROCHA. PARA PIEZA DE MANO. PIEZA.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.219.0068", descripcion: "COLORANTES. REVELADORES DE PLACAS DENTOBACTERIANAS. TABLETAS SIN SABOR. ENVASE CON 100 PIEZAS.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.235.0019", descripcion: "COPAS. PARA PIEZA DE MANO. DE HULE SUAVE BLANCO EN FORMA DE CONO.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.272.0039", descripcion: "CUCHARILLAS. PARA APLICACIÓN TÓPICA DE FLÚOR EN GEL DE VINIL ATÓXICO DESECHABLES. ESTUCHE QUE CONSTA DE: 1 PAR PARA ADOLESCENTES.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.272.0047", descripcion: "CUCHARILLAS. PARA APLICACIÓN TÓPICA DE FLÚOR EN GEL DE VINIL ATÓXICO DESECHABLES. ESTUCHE QUE CONSTA DE: 1 PAR PARA ADULTOS.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.272.0013", descripcion: "CUCHARILLAS. PARA APLICACIÓN TÓPICA DE FLÚOR EN GEL DE VINIL ATÓXICO DESECHABLES. ESTUCHE QUE CONSTA DE: 1 PAR PARA NIÑOS DE 2 A 3 AÑOS.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.272.0021", descripcion: "CUCHARILLAS. PARA APLICACIÓN TÓPICA DE FLÚOR EN GEL DE VINIL ATÓXICO DESECHABLES. ESTUCHE QUE CONSTA DE: 1 PAR PARA NIÑOS DE 4 A 7 AÑOS.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.276.0050", descripcion: "CUÑAS. DE MADERA PARA ESPACIOS INTERDENTARIOS. ENVASE CON 100 PIEZAS.", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "DIQUE DE HULE. 6*6 GROSOR MEDIO  .18 A .23.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.910.0011", descripcion: "EYECTORES. PARA SALIVA DE PLÁSTICO DESECHABLE. ENVASE CON 100 PIEZAS.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.066.1086", descripcion: "FLUORURO DE SODIO EN BARNIZ AL 5%, CONCENTRACIÓN DE 22,600PPM, AUTOPOLIMERIZABLE EN VEHÍCULO DE RESINA MODIFICADO.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.066.0500", descripcion: "FLUORURO DE SODIO. PARA PREVENCIÓN DE CARIES. ACIDULADO AL 2%. EN GEL DE SABOR. ENVASE CON 480 ML.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.066.0112", descripcion: "FLUORURO DE SODIO. SOLUCIÓN INGERIBLE. CADA 100 ML CONTIENEN: FLUORURO DE SODIO EQUIVALENTE A: 248.8 MG DE ION FLÚOR.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0037", descripcion: "FRESA DE CARBURO, PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD, FORMA DE PERA, NO. 331.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0672", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE CARBURO FORMA DE CONO INVERTIDO. NO. 33 1/2.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0656", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE CARBURO, FORMA CILÍNDRICA NO. 556.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0664", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE CARBURO, FORMA CILÍNDRICA NO. 557.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0433", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE CARBURO, FORMA DE CONO INVERTIDO. NO. 37 L.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0631", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE CARBURO, FORMA REDONDA NO. ½.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0649", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE CARBURO, FORMA REDONDA NO. 1.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0409", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE CARBURO, FORMA REDONDA NO. 3.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0417", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE CARBURO, FORMA REDONDA NO. 5.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0466", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE CARBURO, FORMA TRONCO CÓNICO NO.701.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0318", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE CARBURO, NO. 559.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0334", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE CARBURO, NO. 701 L.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0342", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE CARBURO, NO. 702 L.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0011", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE CARBURO, PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD, FORMA DE PERA, NO. 330.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0581", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE DIAMANTE, GRANO GRUESO, FORMA CILÍNDRICA. NO. 009.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0599", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE DIAMANTE, GRANO GRUESO, FORMA CILÍNDRICA. NO. 012.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0540", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE DIAMANTE, GRANO GRUESO, FORMA DE CONO INVERTIDO, NO. 010.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0557", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE DIAMANTE, GRANO GRUESO, FORMA DE CONO INVERTIDO, NO. 012.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0565", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE DIAMANTE, GRANO GRUESO, FORMA DE RUEDA. NO. 035.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0573", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE DIAMANTE, GRANO GRUESO, FORMA DE RUEDA. NO. 042.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0524", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE DIAMANTE, GRANO GRUESO, FORMA REDONDA, NO. 010", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0532", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE DIAMANTE, GRANO GRUESO, FORMA REDONDA, NO. 014.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0615", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE DIAMANTE, PARA TERMINACIÓN DE COMPOSITES FORMA CILÍNDRICA NO. 012", stock: "", minimo: "", caducidad: "" },
      { clave: "060.431.0623", descripcion: "FRESAS DENTALES PARA UTILIZARSE EN LA PIEZA DE MANO DE ALTA VELOCIDAD. DE DIAMANTE, PARA TERMINACIÓN DE COMPOSITES FORMA CILÍNDRICA NO. 018", stock: "", minimo: "", caducidad: "" },
      { clave: "060.811.0078", descripcion: "HILOS. HILO DENTAL DE MONOFILAMENTO.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.203.0439", descripcion: "HILOS. RETRACTOR DE ENCÍAS. DE  ALGODÓN  SECO  Y  SUAVE  IMPREGNADO  CON  SAL  DE ALUMINIO. CALIBRE: MEDIANO.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.811.0060", descripcion: "HILOS. SEDA DENTAL SIN CERA. ENVASE CON ROLLO DE 50 M.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.593.0106", descripcion: "LOSETA. PARA BATIR CEMENTO. DE VIDRIO. TAMAÑO: 8 X 12 X 0.5 CM. PIEZA.", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "MATERIAL ALCASITE DE OBTURACIÓN DEFINITIVO PARA LOS DIENTES POSTERIORES CONTIENE 15 GR DE POLVO Y 4 GR LIQUIDO", stock: "", minimo: "", caducidad: "" },
      { clave: "060.491.0018", descripcion: "PAPELES. INDICADOR DE CONTACTO OCLUSAL. EN TIRAS CON PEGAMENTO EN AMBAS CARAS. BLOCK CON 15 HOJAS.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.749.0703", descripcion: "PASTAS. PARA PROFILAXIS DENTAL. ABRASIVA. CON ABRASIVOS BLANDOS.  ENVASE CON 200 G.", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "PELÍCULA DE PLÁSTICO AUTOADHERIBLE DE  30 CM  DE 280 A  600 MTS DE LARGO.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.749.0836", descripcion: "PULIDOR. SISTEMA DE PULIDO PARA USO DENTAL. KIT. LAS INSTITUCIONES PODRÁN ELEGIR LAS VARIANTES DEL SISTEMA.", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "PUNTAS DE ACABADO Y PULIDO PARA APLICACIONES OCLUSALES E INTERPROXIMALES. TRES GRADOS DE ABRASIÓN Y CUATRO FORMAS DIFERENTES: LLAMA PEQUEÑA, LLAMA GRANDE, COPA Y DISCO.  (PIEDRAS PARA PULIR COMPOSITES)", stock: "", minimo: "", caducidad: "" },
      { clave: "060.791.0106", descripcion: "RESINA. FOTOPOLIMERIZABLE PARA RESTAURACIÓN DE DIENTES ANTERIORES Y POSTERIORES. JERINGA 3.5G.  LAS INSTITUCIONES PODRÁN ELEGIR LAS VARIANTES DE COLOR Y COMPOSICIÓN.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.791.0122", descripcion: "RESINA. FOTOPOLIMERIZABLE PARA RESTAURACIÓN DE DIENTES ANTERIORES Y POSTERIORES. JERINGA 3G.  LAS INSTITUCIONES PODRÁN ELEGIR LAS VARIANTES DE COLOR Y COMPOSICIÓN.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.791.0114", descripcion: "RESINA. FOTOPOLIMERIZABLE PARA RESTAURACIÓN DE DIENTES ANTERIORES Y POSTERIORES. JERINGA 4G.  LAS INSTITUCIONES PODRÁN ELEGIR LAS VARIANTES DE COLOR Y COMPOSICIÓN.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.597.0037", descripcion: "RESINAS. AUTOPOLIMERIZABLES. PARA RESTAURACIÓN DE DIENTES ANTERIORES. EPÓXICAS A BASE DE CUARZO Y AGLUTINANTES. ESTUCHE CON BASE Y CATALIZADOR.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.815.0058", descripcion: "SELLADORES. DE FISURAS Y FOSETAS. ENVASE CON 3 ML DE BOND BASE. ENVASE CON 3 ML DE SELLADOR DE FISURAS. 2 ENVASES CON 3 ML CADA UNO CON BOND CATALIZADOR. JERINGA CON 2 ML DE GEL GRABADOR. 2 PORTAPINCELES. 10 CÁNULAS. 1 BLOCK DE MEZCLA. 5 POZOS DE MEZCLA. 30 PINCELES. 1 INSTRUCTIVO. ESTUCHE.", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "SELLADORES. DE FISURAS Y FOSETAS. KIT CON DOS JERINGAS DE 1.2 ML. INCLUYE UNA JERINGA DE 3ML DE GEL GRABADOR, 20 PUNTAS PARA JERINGA DE SELLADOR Y 10 PUNTAS PARA JERINGA DE GEL GRABADOR.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.833.0197", descripcion: "SOLUCIONES. DE ACETATO DE CLORHEXIDINA AL 10% SUMATRA BENZOICO 20 MG Y ALCOHOL ETÍLICO CBP 1 ML; BARNIZ DE CLORURO DE METILENO POLIURETANO Y ACETATO DE ETILO 1 ML PARA LA PREVENCIÓN DE CARIES DENTAL.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.889.0158", descripcion: "TIRAS. DE CELULOIDE PARA CONFORMAR RESTAURACIONES DE RESINA. ANCHO: 8 A10 MM CALIBRE:", stock: "", minimo: "", caducidad: "" },
      { clave: "060.889.0224", descripcion: "TIRAS. DE LIJA PARA PULIR RESTAURACIONES DE RESINA. EXTRA FINO  MEDIANO Y GRUESO.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.889.0232", descripcion: "TIRAS. DE LIJA PARA PULIR RESTAURACIONES DE RESINA. EXTRA FINO  MEDIANO Y GRUESO.", stock: "", minimo: "", caducidad: "" },
      { clave: "060.889.0208", descripcion: "TIRAS. DE LIJA PARA PULIR RESTAURACIONES DE RESINA. GRUESO Y MEDIANO.", stock: "", minimo: "", caducidad: "" }
    ],
    
    equipo: [
      { clave: "S/C", descripcion: "TINA ULTRASÓNICA PARA ESTOMATOLOGÍA", stock: "", minimo: "", caducidad: "" },
      { clave: "531.385.1080", descripcion: "ESTERILIZADOR CON VAPOR AUTOGENERADO PARA DENTAL Y MAXILOFACIAL", stock: "", minimo: "", caducidad: "" },
      { clave: "531.291.0028", descripcion: "UNIDAD ESTOMATOLÓGICA CON MÓDULO INTEGRADO", stock: "", minimo: "", caducidad: "" },
      { clave: "515.957.0109", descripcion: "UNIDAD ULTRASÓNICA ESTOMATOLÓGICA", stock: "", minimo: "", caducidad: "" },
      { clave: "531.562.0020", descripcion: "LÁMPARA DE FOTOCURADO DE RESINAS Y CEMENTOS FOTOPOLIMERIZABLES", stock: "", minimo: "", caducidad: "" },
      { clave: "531.291.0424", descripcion: "UNIDAD ESTOMATOLÓGICA PORTÁTIL *", stock: "", minimo: "", caducidad: "" },
      { clave: "531.032.0055", descripcion: "MEZCLADOR DE CÁPSULAS DENTALES, TIPO DOSIFICADOR-AMALGAMADOR (MEZCLA CÁPSULAS DE CEMENTO DE IONÓMERO DE VIDRIO Y OTROS CEMENTOS PREDOSIFICADOS EN CÁPSULAS) 4000 RPM", stock: "", minimo: "", caducidad: "" },
      { clave: "531.601.0056", descripcion: "MANDIL EMPLOMADO", stock: "", minimo: "", caducidad: "" },
      { clave: "531.695.0061", descripcion: "PORTA MANDIL", stock: "", minimo: "", caducidad: "" },
      { clave: "531.786.0079", descripcion: "REVELADOR MANUAL DE PLACAS DENTALES", stock: "", minimo: "", caducidad: "" },
      { clave: "531.341.2305", descripcion: "UNIDAD RADIOLÓGICA DENTAL INCLUYE SENSOR DIGITAL INTRAORAL (RADIOVISIÓGRAFO)", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "CÁMARA INTRAORAL DIGITAL.", stock: "", minimo: "", caducidad: "" },
      { clave: "537.480.0042", descripcion: "PIEZA DE MANO DE ALTA VELOCIDAD", stock: "", minimo: "", caducidad: "" },
      { clave: "437.480.0034", descripcion: "PIEZA DE MANO DE BAJA VELOCIDAD CON CONTRAÁNGULO", stock: "", minimo: "", caducidad: "" }
    ],
    mobiliario: [
      { clave: "511.232.0022", descripcion: "CESTO DE PAPELES", stock: "", minimo: "", caducidad: "" },
      { clave: "519.132.0059", descripcion: "BOTE DE CAMPANA", stock: "", minimo: "", caducidad: "" },
      { clave: "511.814.0127", descripcion: "SILLA FIJA ACOJINADA APILABLE", stock: "", minimo: "", caducidad: "" },
      { clave: "515.957.0109", descripcion: "VITRINA DE 75 CM CONTRA MURO", stock: "", minimo: "", caducidad: "" },
      { clave: "511.836.0295", descripcion: "SILLÓN GIRATORIO DE RESPALDO BAJO", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "VITRINA DOBLE ESMALTADA DOS PUERTAS SUPERIORES DE CRISTAL CON CERRADURA Y ENTREPAÑOS DE CRISTAL, DOS CAJONES CENTRALES, DOS PUERTAS DE LÁMINA DE HACER  O INFERIORES CON ENTREPAÑOS", stock: "", minimo: "", caducidad: "" },
      { clave: "511.339.0206", descripcion: "ESCRITORIO CHICO CON PEDESTAL DERECHO", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "CARRO CAJONERO CON RUEDAS", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "MESA ALTA CON RESPALDO Y FREGADERO DE 120 CM", stock: "", minimo: "", caducidad: "" },
      { clave: "511.076.0203", descripcion: "ARCHIVERO DE 4 GAVETAS", stock: "", minimo: "", caducidad: "" },
      { clave: "350.308.0040", descripcion: "DESPCHADOR DE TOALLA DE PAPEL INTERDOBLADA GRANDE", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "LAVABO CONTRA MURO", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "ESPEJO DE PARED", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "ESPEJO INDIVIDUAL PARA PACIENTES", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "AIRE ACONDICIONADO", stock: "", minimo: "", caducidad: "" }
    ],
    bienesInformaticos: [
      { clave: "S/C", descripcion: "COMPUTADORA DE ESCRITORIO MEDIO RENDIMIENTO WINDOWS", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "UNIDAD DE ENERGÍA ININTERRUMPIDA TIPO 3. INTERACTIVO, CAPACIDAD DE 500 VA/300 WATSS.", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "EQUIPO MULTIFUNCIONAL LÁSER MONOCROMÁTICO", stock: "", minimo: "", caducidad: "" }
    ],
    instrumental: [
      { clave: "537.025.0069", descripcion: "ALVEOLOTOMO MEAD PINZA GUBIA LONGITUD 17CM", stock: "", minimo: "", caducidad: "" },
      { clave: "537.065.0052", descripcion: " APLICADOR DE HIDROXIDO DE CALCIO", stock: "", minimo: "", caducidad: "" },
      { clave: "537.079.0015", descripcion: " ARCO YOUNG PORTA DIQUE DE HULE", stock: "", minimo: "", caducidad: "" },
      { clave: "535.137.0035", descripcion: "BISTURÍ QUIRÚRGICO. MANGO N° 3, CON ESCALA", stock: "", minimo: "", caducidad: "" },
      { clave: "537.105.0179", descripcion: "BISTURÍ GOLDMAN FOX. NO.7", stock: "", minimo: "", caducidad: "" },
      { clave: "537.151.0016", descripcion: "BOTA FRESA. BOTA FRESA BRODEN: ADITAMIENTO PARA PIEZA DE MANO DE ALTA VELOCIDAD", stock: "", minimo: "", caducidad: "" },
      { clave: "537.147.0017", descripcion: "BUDINERA DE ACERO INOXIDABLE 25 X 16 CM Y 700 ML.", stock: "", minimo: "", caducidad: "" },
      { clave: "535.260.2154", descripcion: "CUCHARILLA LUCAS DE DOBLE EXTREMO 17CM DE LONGITUD", stock: "", minimo: "", caducidad: "" },
      { clave: "537.251.0098", descripcion: "CURETA. CURETA CK-6 DE DOBLE EXTREMO", stock: "", minimo: "", caducidad: "" },
      { clave: "537.251.0015", descripcion: "CURETA. CURETA MC CALL DERECHA E IZQUIERDA. JUEGO.", stock: "", minimo: "", caducidad: "" },
      { clave: "537.327.0452", descripcion: "ELEVADOR BEIN CON MANGO METALICO RECTO ACANALADO DE 2 O 3 MM ANCHO DE HOJA", stock: "", minimo: "", caducidad: "" },
      { clave: "537.327.0551", descripcion: "ELEVADOR BUCO DENTOMAXILAR. ELEVADOR APICAL FLOHR CON MANGO METALICO CON BRAZO ANGULAR EXTREMO FINO Y CORTO DERECHO", stock: "", minimo: "", caducidad: "" },
      { clave: "537.327.0700", descripcion: "ELEVADOR BUCO DENTOMAXILAR. ELEVADOR APLICAR FLOHR CON MANGO METALICO CON BRAZO ANGULAR EXTREMO FINO Y CORTO IZQUIERDO", stock: "", minimo: "", caducidad: "" },
      { clave: "537.327.2664", descripcion: "ELEVADOR BUCO DENTOMAXILAR. ELEVADOR BEIN CON MANGO METALICO RECTO ACALANADO DE 4MM. ANCHO DE HOJA.", stock: "", minimo: "", caducidad: "" },
      { clave: "537.327.0957", descripcion: "ELEVADOR BUCO DENTOMAXILAR. ELEVADOR TIPO CRYER WHITE DE BANDERA DERECHO CON MANGO METALICO EXTREMO EN ANGULO OBTUSO Y HOJA PEQUEÑA", stock: "", minimo: "", caducidad: "" },
      { clave: "537.327.1104", descripcion: "ELEVADOR BUCO DENTOMAXILAR. ELEVADOR TIPO CRYER WHITE DE BANDERA IZQUIERDO CON MANGO METALICO EXTREMO EN ANGULO OBTUSO Y HOJA PEQUEÑA", stock: "", minimo: "", caducidad: "" },
      { clave: "537.327.1534", descripcion: "ELEVADOR SELDIN DE BANDERA DERECHO MANGO METALICO EXTREMO EN ANGULO RECTO HOJA GRANDE", stock: "", minimo: "", caducidad: "" },
      { clave: "537.327.1609", descripcion: "ELEVADOR SELDIN DE BANDERA IZQUIERDO MANGO METALICO EXTREMO EN ANGULO RECTO HOJA GRANDE", stock: "", minimo: "", caducidad: "" },
      { clave: "537.327.2805", descripcion: "ELEVADOR SELDIN, CON MANGO METÁLICO DE BANDERA, EXTREMO EN ÁNGULO RECTO, CON HOJA PEQUEÑA DERECHA.", stock: "", minimo: "", caducidad: "" },
      { clave: "537.327.2813", descripcion: "ELEVADOR SELDIN, CON MANGO METÁLICO DE BANDERA, EXTREMO EN ÁNGULO RECTO, CON HOJA PEQUEÑA IZQUIERDA", stock: "", minimo: "", caducidad: "" },
      { clave: "537.370.0029", descripcion: "ESPÁTULA METÁLICA DEL No.3 DE DOBLE EXTREMO UNO RECTANGULAR Y OTRO DE PUNTA DE LANZA", stock: "", minimo: "", caducidad: "" },
      { clave: "537.370.0128", descripcion: "ESPÁTULA PARA RESINA, DE PLÁSTICO, CON DOBLE PUNTA DE TRABAJO ", stock: "", minimo: "", caducidad: "" },
      { clave: "537.383.0081", descripcion: "ESPEJO DENTAL ROSCA SENCILLA PLANO SIN AUMENTO NO.5", stock: "", minimo: "", caducidad: "" },
      { clave: "537.397.0168", descripcion: "EXCAVADOR TIPO WHITE No.17", stock: "", minimo: "", caducidad: "" },
      { clave: "537.397.0150", descripcion: "EXCAVADOR TIPO WHITE No.5", stock: "", minimo: "", caducidad: "" },
      { clave: "537.409.0531", descripcion: "EXPLORADOR. EXPLORADOR DE UNA PIEZA CON DOBLE EXTREMO NO.5", stock: "", minimo: "", caducidad: "" },
      { clave: "537.426.0015", descripcion: "FÓRCEPS DENTAL TIPO KLEIN, INFANTIL No. 151 1/2 S ", stock: "", minimo: "", caducidad: "" },
      { clave: "537.426.0189", descripcion: "FÓRCEPS NO. 151", stock: "", minimo: "", caducidad: "" },
      { clave: "537.426.0221", descripcion: "FÓRCEPS NO.101", stock: "", minimo: "", caducidad: "" },
      { clave: "537.426.0197", descripcion: "FÓRCEPS NO.150", stock: "", minimo: "", caducidad: "" },
      { clave: "537.426.0171", descripcion: "FÓRCEPS NO.17", stock: "", minimo: "", caducidad: "" },
      { clave: "537.426.0460", descripcion: "FÓRCEPS NO.53 DERECHO", stock: "", minimo: "", caducidad: "" },
      { clave: "537.426.0270", descripcion: "FÓRCEPS NO.53 IZQUIERDO", stock: "", minimo: "", caducidad: "" },
      { clave: "537.426.0155", descripcion: "FÓRCEPS NO.65", stock: "", minimo: "", caducidad: "" },
      { clave: "537.426.0411", descripcion: "FÓRCEPS NO.69", stock: "", minimo: "", caducidad: "" },
      { clave: "537.426.0205", descripcion: "FÓRCEPS PARA ODONTECTOMÍAS DEL No. 151 B", stock: "", minimo: "", caducidad: "" },
      { clave: "537.426.0726", descripcion: "FÓRCEPS PARA ODONTECTOMÍAS DEL No. 210", stock: "", minimo: "", caducidad: "" },
      { clave: "537.426.0734", descripcion: "FÓRCEPS PARA ODONTECTOMÍAS DEL No. 222", stock: "", minimo: "", caducidad: "" },
      { clave: "537.426.0544", descripcion: "FÓRCEPS PARA ODONTECTOMÍAS TIPO KLEIN, DEL No. 3", stock: "", minimo: "", caducidad: "" },
      { clave: "537.426.0502", descripcion: "FÓRCEPS PARA ODONTECTOMÍAS TIPO KLEIN, DEL No. 6", stock: "", minimo: "", caducidad: "" },
      { clave: "537.426.0023", descripcion: "FÓRCEPS. NO 23.", stock: "", minimo: "", caducidad: "" },
      { clave: "537.547.0019", descripcion: "JERINGA CARPULE CON ADAPTADOR PARA AGUJA DESECHABLE CALIBRE 27 LARFA O CORTA ENTRADA UNIVERSAL O ESTADAR. HENDIDURA PARA INTRODUCIR CARTUCHO DE ANESTESICO DESECHABLE DE 1.8ML DOS ALETAS EN EL CUERPO PARA APOYAR INDICE Y MEDIO. EMBOLO CON ANILLO PARA EL DEDO PULGAR Y ESPIRAL AGUDA EN EL EXTREMO O PUESTO (EN CONTACTO CON LA GOMA DEL CARTUCHO).", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "JUEGO DE GRAPAS", stock: "", minimo: "", caducidad: "" },
      { clave: "S/C", descripcion: "JUEGO DE PROFILAXIS 8 INSTRUMENTOS", stock: "", minimo: "", caducidad: "" },
      { clave: "535.567.0059", descripcion: "LEGRA MEAD MANGO RECTO DOBLE EXTREMO", stock: "", minimo: "", caducidad: "" },
      { clave: "537.583.0105", descripcion: "LIMA MILLER O COLBURN, DOBLE EXTREMO No. 10c O No. 3, C/PUNTAS DE TRABAJO RECTANGULAR Y OVAL, ESTRÍAS HORIZONTALES", stock: "", minimo: "", caducidad: "" },
      { clave: "537.602.0409", descripcion: "MANGO PARA ESPEJO DENTAL METALICO MACIZO ROSCA SENCILLA", stock: "", minimo: "", caducidad: "" },
      { clave: "537.703.9598", descripcion: "PINZA COLLEGE O LONDON-COLLEGE TIPO BAYONETA LONGITUD DE 150 A 155 MM", stock: "", minimo: "", caducidad: "" },
      { clave: "537.703.7493", descripcion: "PINZA AINSWORTH LONGITUD DE 160 A 165 MM", stock: "", minimo: "", caducidad: "" },
      { clave: "537.702.0531", descripcion: "PINZA BREWER PORTA GRAPAS PARA DIQUE DE HULE LONGITUD DE 170 A 175MM", stock: "", minimo: "", caducidad: "" },
      { clave: "531.687.0012", descripcion: "PORTA ABATELENGUAS CON TAPA, DE ACERO INOXIDABLE", stock: "", minimo: "", caducidad: "" },
      { clave: "535.716.0190", descripcion: "PORTA AGUJA FINOCHIETO, LONGITUD 14.6 CM.", stock: "", minimo: "", caducidad: "" },
      { clave: "537.719.0052", descripcion: "PORTA MATRIZ. PORTA MATRIZ PARA BANDA DE CELULOIDE", stock: "", minimo: "", caducidad: "" },
      { clave: "537.720.0018", descripcion: "PORTA SERVILLETAS. MODELO MARTIN O ADAMS CON CADENA", stock: "", minimo: "", caducidad: "" },
      { clave: "060.830.7237", descripcion: "SONDA PARODONTAL. SONDA DE ACERO INOXIDABLE CON UNA PUNTA DE TRABAJO ROMA Y MILIMETRADA DE 1 A 10", stock: "", minimo: "", caducidad: "" },
      { clave: "537.173.2511", descripcion: "SONDA PERIODONTAL W O WHO.", stock: "", minimo: "", caducidad: "" },
      { clave: "535.859.2417", descripcion: "TIJERA MAYO, RECTA LONGITUD DE 170 MM", stock: "", minimo: "", caducidad: "" },
      { clave: "535.859.1898", descripcion: "TIJERA QUINBY, CURVA, HOJAS CORTAS, LONGITUD 12.5 CM", stock: "", minimo: "", caducidad: "" },
      { clave: "535.859.1286", descripcion: "TIJERAS IRIS, CURVA, LONGITUD 12 CM", stock: "", minimo: "", caducidad: "" },
      { clave: "537.860.0018", descripcion: "TIRAPUENTE. TIRAPUENTE MILLER CON TRES PUNTAS DIFERENTES", stock: "", minimo: "", caducidad: "" },
      { clave: "513.887.0059", descripcion: "TORUNDERA CON TAPA, DE ACERO INOXIDABLE DE 250 ML. DE CAPACIDAD", stock: "", minimo: "", caducidad: "" }
    ]
    
  };
  // --------------------------------------------------------------------

  // DOM
  const selCategoria = document.getElementById("categoria");
  const btnSiguiente = document.getElementById("btnSiguiente");
  const btnRegresar = document.getElementById("btnRegresar1");
  const btnAgregar = document.getElementById("btnAgregarFila");
  const tbody = document.querySelector("#tablaInsumos tbody");
  const tituloCategoria = document.getElementById("tituloCategoria");
  const tabla = document.getElementById("tablaInsumos");
  const btnEnviar = document.getElementById("btnEnviarInsumos");

  // Asegurar encabezado: agregar columna "Observaciones" después de "Días restantes"
  function ensureObservacionesHeader() {
    if (!tabla) return;
    let thead = tabla.querySelector("thead");
    if (!thead) {
      // si no hay thead, crear uno e intentar copiar primera fila como header si existe
      thead = document.createElement("thead");
      tabla.insertBefore(thead, tabla.firstChild);
      const headerRow = document.createElement("tr");
      const defaultHeaders = ["No.", "Clave", "Descripción", "Stock", "Mínimo", "Estado", "Caducidad", "Días restantes", "Observaciones"];
      defaultHeaders.forEach(h => {
        const th = document.createElement("th");
        th.textContent = h;
        headerRow.appendChild(th);
      });
      thead.appendChild(headerRow);
      return;
    }
    const ths = Array.from(thead.querySelectorAll("th"));
    // si ya existe Observaciones, salir
    if (ths.some(t => (t.textContent || "").trim().toLowerCase() === "observaciones")) return;

    // intentar insertar después del th que contenga "Días restantes"
    let insertIndex = ths.findIndex(t => ((t.textContent || "").trim().toLowerCase().includes("días") && (t.textContent || "").trim().toLowerCase().includes("restantes")));
    if (insertIndex === -1) {
      // buscar por "Días" o "dias"
      insertIndex = ths.findIndex(t => ((t.textContent || "").trim().toLowerCase().includes("días") || (t.textContent || "").trim().toLowerCase().includes("dias")));
    }
    const thObs = document.createElement("th");
    thObs.textContent = "Observaciones";
    if (insertIndex >= 0) {
      // insertar después de ese índice
      const ref = ths[insertIndex].nextSibling;
      if (ref) thead.querySelector("tr").insertBefore(thObs, ref);
      else thead.querySelector("tr").appendChild(thObs);
    } else {
      // si no encontramos referencia, simplemente añadir al final
      thead.querySelector("tr").appendChild(thObs);
    }
  }
  ensureObservacionesHeader();

  // ---------------- Navegación ----------------
  // usar onclick para evitar handlers duplicados antiguos
  btnSiguiente.onclick = (ev) => {
    ev && ev.preventDefault();
    const cat = selCategoria.value;
    if (!cat) return alert("Selecciona una categoría.");
    categoriaActiva = cat;
    tituloCategoria.textContent = `Formulario de ${selCategoria.options[selCategoria.selectedIndex].text}`;
    document.getElementById("page1").classList.remove("activo");
    document.getElementById("page1").classList.add("oculto");
    document.getElementById("page2").classList.remove("oculto");
    document.getElementById("page2").classList.add("activo");
    limpiarTabla();
    agregarFila(); // fila inicial
  };

  btnRegresar.onclick = (ev) => {
    ev && ev.preventDefault();
    document.getElementById("page2").classList.remove("activo");
    document.getElementById("page2").classList.add("oculto");
    document.getElementById("page1").classList.remove("oculto");
    document.getElementById("page1").classList.add("activo");
    limpiarTabla();
    categoriaActiva = null;
    selCategoria.value = "";
  };

  function limpiarTabla() {
    if (tbody) tbody.innerHTML = "";
    filaContador = 0;
  }

  // Evitar múltiples agregados: usaremos onclick + bloqueo temporal del botón
  btnAgregar.onclick = (ev) => {
    ev && ev.preventDefault();
    if (!categoriaActiva) {
      alert("Selecciona primero una categoría.");
      return;
    }
    const now = Date.now();
    if (now - lastAddTime < 250) return; // bloqueo rápido
    lastAddTime = now;
    // bloquear botón unos ms para evitar dobles envíos
    btnAgregar.disabled = true;
    setTimeout(() => { btnAgregar.disabled = false; }, 300);
    agregarFila();
  };

  function agregarFila() {
    filaContador++;
    const tr = document.createElement("tr");

    // No.
    const tdNo = document.createElement("td");
    tdNo.textContent = filaContador;
    tr.appendChild(tdNo);

    // Clave (select)
    const tdClave = document.createElement("td");
    const select = document.createElement("select");
    const optDefault = document.createElement("option");
    optDefault.value = "";
    optDefault.textContent = "--Seleccione--";
    select.appendChild(optDefault);

    if (catalogo[categoriaActiva] && catalogo[categoriaActiva].length > 0) {
      catalogo[categoriaActiva].forEach((p, idx) => {
        const o = document.createElement("option");
        // guardamos índice para identificar exactamente el producto (evita confusiones con S/C)
        o.value = `${p.clave}||${idx}`;
        o.textContent = p.clave;
        o.dataset.descripcion = p.descripcion || "";
        o.dataset.idx = String(idx);
        select.appendChild(o);
      });
    }
    tdClave.appendChild(select);
    tr.appendChild(tdClave);

    // Descripción (editable + datalist)
    const tdDesc = document.createElement("td");
    const inputDesc = document.createElement("input");
    inputDesc.type = "text";
    inputDesc.placeholder = "Escribe descripción o selecciona sugerencia";
    inputDesc.readOnly = false;
    inputDesc.tabIndex = 0;

    const datalistId = `datalist-desc-${Date.now()}-${Math.random().toString(36).slice(2,6)}`;
    const dl = document.createElement("datalist");
    dl.id = datalistId;
    if (catalogo[categoriaActiva]) {
      catalogo[categoriaActiva].forEach(p => {
        const opt = document.createElement("option");
        opt.value = p.descripcion;
        dl.appendChild(opt);
      });
    }
    inputDesc.setAttribute("list", datalistId);
    tdDesc.appendChild(inputDesc);
    tdDesc.appendChild(dl);
    tr.appendChild(tdDesc);

    // Stock
    const tdStock = document.createElement("td");
    const inputStock = document.createElement("input");
    inputStock.type = "number";
    inputStock.min = 0;
    tdStock.appendChild(inputStock);
    tr.appendChild(tdStock);

    // Mínimo
    const tdMin = document.createElement("td");
    const inputMin = document.createElement("input");
    inputMin.type = "number";
    inputMin.min = 0;
    tdMin.appendChild(inputMin);
    tr.appendChild(tdMin);

    // Estado
    const tdEstado = document.createElement("td");
    const spanEstado = document.createElement("span");
    tdEstado.appendChild(spanEstado);
    tr.appendChild(tdEstado);

    // Caducidad
    const tdCad = document.createElement("td");
    const inputCad = document.createElement("input");
    inputCad.type = "date";
    tdCad.appendChild(inputCad);
    tr.appendChild(tdCad);

    // Días restantes
    const tdDias = document.createElement("td");
    const inputDias = document.createElement("input");
    inputDias.type = "text";
    inputDias.readOnly = true;
    inputDias.value = "";
    tdDias.appendChild(inputDias);
    tr.appendChild(tdDias);

    // Observaciones (editable) - columna nueva
    const tdObs = document.createElement("td");
    const inputObs = document.createElement("input");
    inputObs.type = "text";
    inputObs.placeholder = "Observaciones (texto libre)";
    inputObs.readOnly = false;
    tdObs.appendChild(inputObs);
    tr.appendChild(tdObs);

    tbody.appendChild(tr);

    // ------------------ Lógica de autocompletado ------------------

    function fillProduct(producto) {
      if (!producto) return;
      inputDesc.value = producto.descripcion || inputDesc.value;
      inputStock.value = producto.stock || "";
      inputMin.value = producto.minimo || "";
      inputCad.value = producto.caducidad || "";
      const lista = catalogo[categoriaActiva] || [];
      const idx = lista.indexOf(producto);
      if (idx >= 0) {
        for (let i = 0; i < select.options.length; i++) {
          const opt = select.options[i];
          if (opt.dataset && ('idx' in opt.dataset) && parseInt(opt.dataset.idx, 10) === idx) {
            select.value = opt.value;
            break;
          }
        }
      } else {
        select.value = `${producto.clave}||0`;
      }
      actualizarFila(tr);
    }

    // inputDesc: solo autocompleta en coincidencia exacta al escribir; datalist mostrará sugerencias
    inputDesc.addEventListener("input", () => {
      const v = (inputDesc.value || "").trim();
      if (!v) {
        actualizarFila(tr);
        return;
      }
      const lista = catalogo[categoriaActiva] || [];
      const vLower = v.toLowerCase();
      const productoExact = lista.find(p => p.descripcion && p.descripcion.trim().toLowerCase() === vLower);
      if (productoExact) {
        fillProduct(productoExact);
        return;
      }
      // si no exacto, no sobrescribimos (dejamos al usuario editar)
      actualizarFila(tr);
    });

    // Enter/Tab confirma y puede autocompletar si hay coincidencias claras
    inputDesc.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter" || ev.key === "Tab") {
        const v = (inputDesc.value || "").trim();
        if (!v) return;
        const lista = catalogo[categoriaActiva] || [];
        const vLower = v.toLowerCase();

        const matchesStarts = lista.filter(p => p.descripcion && p.descripcion.trim().toLowerCase().startsWith(vLower));
        const matchesContains = lista.filter(p => p.descripcion && p.descripcion.trim().toLowerCase().includes(vLower));

        let producto = null;
        if (matchesStarts.length === 1) producto = matchesStarts[0];
        else if (matchesStarts.length > 1) producto = matchesStarts[0];
        else if (matchesContains.length === 1) producto = matchesContains[0];
        else if (matchesContains.length > 1 && matchesContains[0].descripcion.trim().toLowerCase().startsWith(vLower)) producto = matchesContains[0];

        if (producto) {
          // si usamos Tab, no prevenimos para permitir cambio de foco si el usuario quiere
          if (ev.key === "Enter") ev.preventDefault();
          fillProduct(producto);
        }
      }
    });

    // select -> autocompleta
    select.addEventListener("change", () => {
      const selectedOption = select.selectedOptions[0];
      let producto = null;
      if (selectedOption && selectedOption.dataset && ('idx' in selectedOption.dataset)) {
        const idx = parseInt(selectedOption.dataset.idx, 10);
        producto = (catalogo[categoriaActiva] || [])[idx] || null;
      }
      if (!producto) {
        const claveSimple = select.value ? select.value.split("||")[0] : "";
        producto = (catalogo[categoriaActiva] || []).find(p => p.clave === claveSimple) || null;
      }
      if (producto) fillProduct(producto);
      setTimeout(() => { try { inputDesc.focus(); } catch(e) {} }, 0);
    });

    // stock / min / cad
    inputStock.addEventListener("input", () => {
      if (inputStock.value === "") return actualizarFila(tr);
      let v = parseInt(inputStock.value, 10);
      if (isNaN(v) || v < 0) v = 0;
      inputStock.value = v;
      actualizarFila(tr);
    });

    inputMin.addEventListener("input", () => {
      if (inputMin.value === "") return actualizarFila(tr);
      let v = parseInt(inputMin.value, 10);
      if (isNaN(v) || v < 0) v = 0;
      inputMin.value = v;
      actualizarFila(tr);
    });

    inputCad.addEventListener("change", () => actualizarFila(tr));
  }

  // Actualizar fila (estado + semáforo)
  function actualizarFila(tr) {
    const inputStock = tr.cells[3].querySelector("input");
    const inputMin = tr.cells[4].querySelector("input");
    const inputCad = tr.cells[6].querySelector("input");
    const inputDias = tr.cells[7].querySelector("input");
    const estadoSpan = tr.cells[5].querySelector("span");

    const stockVal = inputStock.value === "" ? null : Math.max(0, parseInt(inputStock.value || 0, 10));
    const minVal = inputMin.value === "" ? 0 : Math.max(0, parseInt(inputMin.value || 0, 10));
    if (stockVal === null) {
      estadoSpan.textContent = "";
    } else {
      estadoSpan.textContent = (stockVal < minVal) ? "Bajo stock" : "Stock suficiente";
    }

    tr.classList.remove("expired", "warning-expiry", "valid-expiry");
    inputDias.value = "";

    if (inputCad.value) {
      const hoy = new Date();
      hoy.setHours(0,0,0,0);
      const cad = new Date(inputCad.value);
      const diffMs = cad - hoy;
      const diasRest = Math.ceil(diffMs / (1000*60*60*24));
      inputDias.value = diasRest < 0 ? "Caducado" : diasRest;

      let meses = (cad.getFullYear() - hoy.getFullYear()) * 12 + (cad.getMonth() - hoy.getMonth());
      if (cad.getDate() < hoy.getDate()) meses -= 1;

      if (meses < 0) {
        tr.classList.add("expired");
      } else if (meses < 6) {
        tr.classList.add("expired");
      } else if (meses <= 12) {
        tr.classList.add("warning-expiry");
      } else {
        tr.classList.add("valid-expiry");
      }
    }
  }

  // Botón enviar: recopila datos, limpia la tabla y regresa a página inicial
  if (btnEnviar) {
    btnEnviar.textContent = "Enviar";
    btnEnviar.onclick = (ev) => {
      ev && ev.preventDefault();
      const datos = [];
      for (const row of tbody.rows) {
        const select = row.cells[1].querySelector("select");
        const raw = select ? select.value : "";
        const claveReal = raw ? raw.split("||")[0] : "";
        const descripcion = (row.cells[2].querySelector("input").value || "").trim();
        // OBS: observaciones está en la última celda: índice 8 si mantienes columnas como en el script
        const obsCellIndex = row.cells.length - 1;
        const observaciones = (row.cells[obsCellIndex].querySelector("input").value || "").trim();
        if (!claveReal && !descripcion && !observaciones) continue;
        datos.push({
          clave: claveReal,
          descripcion,
          stock: row.cells[3].querySelector("input").value,
          minimo: row.cells[4].querySelector("input").value,
          caducidad: row.cells[6].querySelector("input").value,
          diasRestantes: row.cells[7].querySelector("input").value,
          observaciones
        });
      }
      console.log("Datos a enviar:", datos);
      alert("Datos listos (ver consola).");

      // limpieza y reset
      limpiarTabla();
      categoriaActiva = null;
      selCategoria.value = "";
      // volver a la página 1
      document.getElementById("page2").classList.remove("activo");
      document.getElementById("page2").classList.add("oculto");
      document.getElementById("page1").classList.remove("oculto");
      document.getElementById("page1").classList.add("activo");
    };
  }
});
