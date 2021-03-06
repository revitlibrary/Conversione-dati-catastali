Friend Class DatiTabelle
    Friend DatiTABELLE As New ArrayList

    Public Sub New()

        'DatiTABELLE.Add("CREATE TABLE [CTTITOLI] ( CODICE VARCHAR(3), TITOLO VARCHAR(53) )")
        'DatiTABELLE.Add("CREATE TABLE [CTQUALIT] ( CODICE Double, COD_QUALIT VARCHAR(18), QUALITA VARCHAR(12) )")
        'DatiTABELLE.Add("CREATE TABLE [CTCOMCAT] ( COD_CAT VARCHAR(5), CODICE VARCHAR(4), COMUNE VARCHAR(65), SEZIONE VARCHAR(1) )")
        'DatiTABELLE.Add("CREATE TABLE [CUCOMCAT] ( COD_CAT VARCHAR(5), CODICE VARCHAR(4), COMUNE VARCHAR(65), SEZIONE VARCHAR(1) )")
        'DatiTABELLE.Add("CREATE TABLE [CUTIPNOT] ( DESCRIZION VARCHAR(35), TIPO_NOTA VARCHAR(1) )")
        'DatiTABELLE.Add("CREATE TABLE [CUCODTOP] ( CODICE Double, TOPONIMO VARCHAR(30) )")
        'DatiTABELLE.Add("CREATE TABLE [CTCOMNAZ] ( CODICE VARCHAR(4), COMUNE VARCHAR(54), PROVINCIA VARCHAR(2) )")
        'DatiTABELLE.Add("CREATE TABLE [CUAPPCOM] ( COD_SEZ VARCHAR(5), COMUNE VARCHAR(65) )")

        DatiTABELLE.Add("CREATE TABLE [CTFISICA] ( CODICE VARCHAR(4), SEZIONE VARCHAR(1), SOGGETTO Double UNIQUE, TIPO_SOG VARCHAR(1), COGNOME VARCHAR(50), NOME VARCHAR(50), SESSO VARCHAR(1), DATA VARCHAR(10), LUOGO VARCHAR(4), CODFISCALE VARCHAR(16), SUPPLEMENT VARCHAR(100))")

        DatiTABELLE.Add("CREATE TABLE [CTNONFIS] ( CODICE VARCHAR(4), SEZIONE VARCHAR(1),SOGGETTO Double UNIQUE, TIPO_SOG VARCHAR(1), DENOMINAZ VARCHAR(150), SEDE VARCHAR(4), CODFISCALE VARCHAR(11) )")

        DatiTABELLE.Add("CREATE TABLE [CTTITOLA] ( CODICE VARCHAR(4),SEZIONE VARCHAR(1), SOGGETTO Double, TIPO_SOG VARCHAR(1), IMMOBILE Double,TIPO_IMM VARCHAR(1), DIRITTO VARCHAR(3),TITOLO VARCHAR(200),NUMERATORE Double, DENOMINATO Double" &
                                                 ",REGIME VARCHAR(1), RIF_REGIME Double,GEN_VALIDA VARCHAR(10), GEN_NOTA VARCHAR(1), GEN_NUMERO VARCHAR(6),GEN_PROGRE VARCHAR(3), GEN_ANNO VARCHAR(4),GEN_REGIST VARCHAR(10)" &
                                                 ",PARTITA VARCHAR(7), CON_VALIDA VARCHAR(10),CON_NOTA VARCHAR(1), CON_NUMERO VARCHAR(6),CON_PROGRE VARCHAR(3),CON_ANNO VARCHAR(4)" &
                                                 ",CON_REGIST VARCHAR(10),MUTAZ_INIZ Double,MUTAZ_FINE Double,IDENTIFICA Double,FLAG_IMPOR VARCHAR(6) )")


        DatiTABELLE.Add("CREATE TABLE [CUINDIRI] ( CODICE VARCHAR(4),SEZIONE VARCHAR(1),IMMOBILE Double,TIPO_IMM VARCHAR(1),PROGRESSIV Double,TOPONIMO Double,INDIRIZZO VARCHAR(50),CIVICO1 VARCHAR(6),CIVICO2 VARCHAR(6),CIVICO3 VARCHAR(6), FLAG_IMPOR VARCHAR(3))")

        DatiTABELLE.Add("CREATE TABLE [CUARCUIU] ( ANNOTAZION VARCHAR(200), CATEGORIA VARCHAR(3), CLASSE VARCHAR(2), CODICE VARCHAR(4), CON_ANNO VARCHAR(4)" &
                                                ", CON_EFF VARCHAR(10), CON_NUMERO VARCHAR(6), CON_PROGRE VARCHAR(3), CON_REGIST VARCHAR(10), CON_TIPO VARCHAR(1), CONSISTENZ VARCHAR(7)" &
                                                ", EDIFICIO VARCHAR(2), GEN_ANNO VARCHAR(4), GEN_EFF VARCHAR(10), GEN_NUMERO VARCHAR(6), GEN_PROGRE VARCHAR(3), GEN_REGIST VARCHAR(10), GEN_TIPO VARCHAR(1)" &
                                                ", IMMOBILE Double, INTERNO_1 VARCHAR(3), INTERNO_2 VARCHAR(3), LOTTO VARCHAR(2), MUTAZ_FINE Double, MUTAZ_INIZ Double, PARTITA VARCHAR(7)" &
                                                ", PIANO_1 VARCHAR(4), PIANO_2 VARCHAR(4), PIANO_3 VARCHAR(4), PIANO_4 VARCHAR(4), PROGRESSIV Double, PROT_NOTIF VARCHAR(18), RENDITA_E VARCHAR(18), RENDITA_L VARCHAR(15)" &
                                                ", SCALA VARCHAR(2), SEZIONE VARCHAR(1), SUPERFICIE VARCHAR(5), TIPO_IMM VARCHAR(1), ZONA VARCHAR(3) )")

        DatiTABELLE.Add("CREATE TABLE [CUIDENTI] ( CODICE VARCHAR(4),SEZIONE VARCHAR(1),IMMOBILE Double,TIPO_IMM VARCHAR(1),PROGRESSIV Double,SEZ_URBANA VARCHAR(3),FOGLIO VARCHAR(4),PARTICELLA VARCHAR(5), DENOMINATO Double, " &
                                                           "SUBALTERNO VARCHAR(4),EDIFICIALE VARCHAR(1),FLAG_IMPOR VARCHAR(6),[FP] VARCHAR(11),[FP_INDEX] Integer)")

        DatiTABELLE.Add("CREATE TABLE [CUUTILIT] ( CODICE VARCHAR(4), DENOMINATO Double, FLAG_IMPOR VARCHAR(1), FOGLIO VARCHAR(4), IMMOBILE Double, PARTICELLA VARCHAR(5), PROGRESSIV Double, SEZ_URBANA VARCHAR(3), SEZIONE VARCHAR(1), SUBALTERNO VARCHAR(4), TIPO_IMM VARCHAR(1) )")

        DatiTABELLE.Add("CREATE TABLE [CURISERV] ( CODICE VARCHAR(4), FLAG_IMPOR VARCHAR(1), IMMOBILE Double, ISCRIZIONE VARCHAR(7), PROGRESSIV Double, RISERVA VARCHAR(1), SEZIONE VARCHAR(1), TIPO_IMM VARCHAR(1) )")

        DatiTABELLE.Add("CREATE TABLE [CTPARTIC] ( CODICE VARCHAR(4),SEZIONE VARCHAR(1),IMMOBILE Double,TIPO_IMM VARCHAR(1),PROGRESSIV Double,FOGLIO Double,PARTICELLA VARCHAR(5),DENOMINATO Double,SUBALTERNO VARCHAR(4)," &
                                                            "EDIFICIALE VARCHAR(1),QUALITA Double,[CLASSE] VARCHAR(2),ETTARI Double,[ARE] Double,[CENTIARE] Double,FLAG_REDD VARCHAR(1),FLAG_PORZ VARCHAR(1),FLAG_DEDUZ VARCHAR(1)," &
                                                            "DOMINIC_L VARCHAR(9),AGRARIO_L VARCHAR(8),DOMINIC_E VARCHAR(12),AGRARIO_E VARCHAR(11),GEN_EFF VARCHAR(10),GEN_REGIST VARCHAR(10),GEN_TIPO VARCHAR(1),GEN_NUMERO VARCHAR(6),GEN_PROGRE VARCHAR(3)," &
                                                            "GEN_ANNO Double,CON_EFF VARCHAR(10),CON_REGIST VARCHAR(10),CON_TIPO VARCHAR(1),CON_NUMERO VARCHAR(6),CON_PROGRE VARCHAR(3),CON_ANNO Double,PARTITA VARCHAR(7),ANNOTAZION VARCHAR(200),MUTAZ_INIZ Double,MUTAZ_FINE Double)")

        DatiTABELLE.Add("CREATE TABLE [CTDEDUZI] ( CODICE VARCHAR(4),SEZIONE VARCHAR(1),IMMOBILE Double,TIPO_IMM VARCHAR(1),PROGRESSIV Double,DEDUZIONE VARCHAR(6),FLAG_IMPOR VARCHAR(1))")

        DatiTABELLE.Add("CREATE TABLE [CTPORZIO] ( CODICE VARCHAR(4), SEZIONE VARCHAR(1), IMMOBILE Double, TIPO_IMM VARCHAR(1), PROGRESSIV Double, PORZIONE VARCHAR(2), QUALITA Double,  [CLASSE] VARCHAR(2),ETTARI Double, [ARE] Double, [CENTIARE] Double, FLAG_IMPOR VARCHAR(6))")

        DatiTABELLE.Add("CREATE TABLE [CTRISERV] ( CODICE VARCHAR(4), SEZIONE VARCHAR(1), IMMOBILE Double, TIPO_IMM VARCHAR(1), PROGRESSIV Double, RISERVA VARCHAR(1), ISCRIZIONE VARCHAR(6), FLAG_IMPOR VARCHAR(6))")

        DatiTABELLE.Add("CREATE TABLE [CUAPPUIU] ( ANNOTAZION VARCHAR(200), CATEGORIA VARCHAR(3), CLASSE VARCHAR(2), CODICE VARCHAR(4), CON_ANNO VARCHAR(4), CON_EFF VARCHAR(10),
                                                   CON_NUMERO VARCHAR(6), CON_PROGRE VARCHAR(3), CON_REGIST VARCHAR(10), CON_TIPO VARCHAR(1), CONSISTENZ VARCHAR(7), 
                                                   EDIFICIO VARCHAR(2), GEN_ANNO VARCHAR(4), GEN_EFF VARCHAR(10), GEN_NUMERO VARCHAR(6), GEN_PROGRE VARCHAR(3), GEN_REGIST VARCHAR(10), 
                                                   GEN_TIPO VARCHAR(1), IMMOBILE Double, INTERNO_1 VARCHAR(3), INTERNO_2 VARCHAR(3), LOTTO VARCHAR(2), MUTAZ_FINE Double, 
                                                   MUTAZ_INIZ Double, PARTITA VARCHAR(7), PIANO_1 VARCHAR(4), PIANO_2 VARCHAR(4), PIANO_3 VARCHAR(4), PIANO_4 VARCHAR(4), 
                                                   PROGRESSIV Double, PROT_NOTIF VARCHAR(18), RENDITA_E VARCHAR(18), RENDITA_L VARCHAR(15), SCALA VARCHAR(2), SEZIONE VARCHAR(1),
                                                   SUPERFICIE VARCHAR(5), TIPO_IMM VARCHAR(1), ZONA VARCHAR(3) )")

        DatiTABELLE.Add("CREATE TABLE [CUCOMUNI] ( CODICE VARCHAR(4), COMUNE VARCHAR(65), DATA_ESTRA VARCHAR(10), DATA_IMPOR VARCHAR(10), SEZIONE VARCHAR(1) )")

        DatiTABELLE.Add("CREATE TABLE [CUGRUPPI] ( DESCRIZION VARCHAR(170), GRUPPO VARCHAR(4), REND_EURO Double, RENDITA Double, UIU Double, VANI Double )")

        DatiTABELLE.Add("CREATE TABLE [DATACREA] ( COMUNE VARCHAR(5), DATARICHIESTA VARCHAR(10), DATAELABORAZIONE VARCHAR(10), TipoEstrazione VARCHAR(100), NumRecord Double, PrmArchivio VARCHAR(10))")

        DatiTABELLE.Add("CREATE TABLE [EDIFICI] (FOGLIO VARCHAR(4), PARTICELLA VARCHAR(5),SUBALTERNO VARCHAR(4),CATEGORIA VARCHAR(3),[CLASSE] VARCHAR(2),CONSISTENZ VARCHAR(7),SUPERFICIE VARCHAR(5)," &
                        "SOGGETTO Double,IMMOBILE Double,DIRITTO VARCHAR(3),RAGIONE VARCHAR(100),DATA VARCHAR(10),LUOGO VARCHAR(4),CODFISCALE VARCHAR(16),COMUNE VARCHAR(65),NOME_FILE VARCHAR(255),NOME_FILE_2 VARCHAR(255)," &
                        "RENDITA_E VARCHAR(18),PIANO_1 VARCHAR(4),PIANO_2 VARCHAR(4),PIANO_3 VARCHAR(4),PIANO_4 VARCHAR(4),[CHECK] VARCHAR(1),CATASTO VARCHAR(7),COD_TIPOS Double,TIPOLOGIA VARCHAR(255),FOGLIO_PARTICELLA VARCHAR(15))")

        DatiTABELLE.Add("CREATE TABLE [TERRENI] (FOGLIO VARCHAR(4),PARTICELLA VARCHAR(5),SUBALTERNO VARCHAR(4),RAGIONE VARCHAR(120),DATA VARCHAR(10),LUOGO VARCHAR(120),COMUNE VARCHAR(80),CODFISCALE VARCHAR(20)," &
                        "EDIFICIALE VARCHAR(1),QUALITA Double,QUALITA_TESTO VARCHAR(50),CLASSE VARCHAR(2),[AREA] Double,[ETTARI] Double,[ARE] Double,[CENTIARE] Double,DOMINIC_E VARCHAR(20),AGRARIO_E VARCHAR(20)," &
                        "DIRITTO_PROPRIETA VARCHAR(50),SOGGETTO Double,IMMOBILE Double,[CHECK] VARCHAR(1),CATASTO VARCHAR(10))")

        DatiTABELLE.Add("CREATE TABLE [CATASTINI] ([INDEX] double,FOGLIO VARCHAR(4),PARTICELLA VARCHAR(5),SUBALTERNO VARCHAR(4),[nome_files] VARCHAR(255) )")

    End Sub
End Class

