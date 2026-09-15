-- =====================================================================
-- Cicli a saldo zero entro il 31/08/2016 - approccio in due fasi
-- Db2 for i 7.4 - MOV_EOR_LAV ~30M righe
--
-- Idea: NON calcolare il saldo progressivo su 30M righe.
-- Prima si individuano con una GROUP BY (economica, hash, nessun sort
-- ordinato) le sole coppie titolo+intestatario candidate; poi si applica
-- la window solo a quelle.
-- =====================================================================

-- ---------------------------------------------------------------------
-- INDICE - indispensabile: senza, la window ordina 30M righe in *TEMP
-- ---------------------------------------------------------------------
CREATE INDEX MOV_EOR_LAV_IX1
    ON MOV_EOR_LAV (CHIAVE_TITOLO,
                    CLIENTE_INTESTATARIO_XF,
                    DATA_OPERAZIONE,
                    NUMERO_OPERAZIONE);

-- ---------------------------------------------------------------------
-- FASE 1 - Coppie il cui saldo TOTALE entro la data e' zero.
--          Semplice aggregazione: nessun ORDER BY, nessun framing.
--          Se la definizione richiesta e' "posizione chiusa al 31/08/2016",
--          questa fase da' gia' l'elenco completo delle coppie.
-- ---------------------------------------------------------------------
DROP TABLE IF EXISTS COPPIE_ZERO;

CREATE TABLE COPPIE_ZERO AS (
    SELECT CHIAVE_TITOLO,
           CLIENTE_INTESTATARIO_XF,
           COUNT(*)                AS N_MOVIMENTI,
           SUM(QTA_CON_SEGNO)      AS SALDO_FINALE
      FROM MOV_EOR_LAV
     GROUP BY CHIAVE_TITOLO, CLIENTE_INTESTATARIO_XF
    HAVING SUM(QTA_CON_SEGNO) = 0
) WITH DATA;

CREATE UNIQUE INDEX COPPIE_ZERO_IX1
    ON COPPIE_ZERO (CHIAVE_TITOLO, CLIENTE_INTESTATARIO_XF);

-- Quante righe restano da elaborare nella fase 2?
SELECT COUNT(*) AS COPPIE, SUM(N_MOVIMENTI) AS RIGHE FROM COPPIE_ZERO;

-- ---------------------------------------------------------------------
-- FASE 2 - Saldo progressivo SOLO sulle coppie superstiti.
--          Colonne esplicite (niente t.*): meno byte scritti.
--          Serve solo se interessa anche il dettaglio riga per riga
--          o l'azzeramento INTERMEDIO (vedi nota in fondo).
-- ---------------------------------------------------------------------
DROP TABLE IF EXISTS MOVIMENTI_SALDO;

CREATE TABLE MOVIMENTI_SALDO AS (
    SELECT t.NUMERO_OPERAZIONE,
           t.CHIAVE_TITOLO,
           t.CODICE_ISIN,
           t.CLIENTE_INTESTATARIO_XF,
           t.DATA_OPERAZIONE,
           t.SEGNO_OPERAZIONE,
           t.QUANTITA,
           SUM(t.QTA_CON_SEGNO) OVER (
               PARTITION BY t.CHIAVE_TITOLO, t.CLIENTE_INTESTATARIO_XF
               ORDER BY t.DATA_OPERAZIONE, t.NUMERO_OPERAZIONE
               ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW
           ) AS SALDO_PROGRESSIVO,
           ROW_NUMBER() OVER (
               PARTITION BY t.CHIAVE_TITOLO, t.CLIENTE_INTESTATARIO_XF
               ORDER BY t.DATA_OPERAZIONE, t.NUMERO_OPERAZIONE
           ) AS PROG
      FROM MOV_EOR_LAV t
     WHERE EXISTS (SELECT 1
                     FROM COPPIE_ZERO z
                    WHERE z.CHIAVE_TITOLO           = t.CHIAVE_TITOLO
                      AND z.CLIENTE_INTESTATARIO_XF = t.CLIENTE_INTESTATARIO_XF)
) WITH DATA;

-- ---------------------------------------------------------------------
-- NOTA - azzeramento INTERMEDIO
-- La fase 1 tiene solo le coppie con saldo finale zero. Una coppia che si
-- azzera a meta' periodo e poi riapre posizione resta esclusa.
-- Se servono anche quelle, la fase 1 va cambiata in un filtro piu' debole
-- (es. solo coppie con almeno un carico e almeno uno scarico):
--
--   HAVING MIN(QTA_CON_SEGNO) < 0 AND MAX(QTA_CON_SEGNO) > 0
--
-- che riduce molto meno, ma non perde cicli.
-- ---------------------------------------------------------------------
