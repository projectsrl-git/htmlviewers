-- =====================================================================
-- Cicli a saldo zero entro il 31/08/2016 - versione ottimizzata
-- Db2 for i 7.4 - MOVIMENTI_TITOLI_EOR materializzata come tabella (~20M righe)
-- =====================================================================

-- ---------------------------------------------------------------------
-- STEP 1 - Tabella di lavoro stretta, gia' filtrata e con colonne precalcolate
--          (chiave titolo, quantita' con segno): niente espressioni a runtime,
--          niente filtro G-GESTITO / data da rivalutare a ogni esecuzione.
-- ---------------------------------------------------------------------
CREATE TABLE MOV_EOR_LAV AS (
    SELECT NUMERO_OPERAZIONE,
           CODICE_ISIN,
           CODICE_INTERNO_TITOLO,
           CAST(COALESCE(NULLIF(TRIM(CODICE_ISIN), ''),
                         'INT:' CONCAT TRIM(CODICE_INTERNO_TITOLO))
                AS VARCHAR(20))                           AS CHIAVE_TITOLO,
           CLIENTE_INTESTATARIO_XF,
           DATA_OPERAZIONE,                               -- VARCHAR 'AAAA/MM/GG'
           SEGNO_OPERAZIONE,
           CONTO_AMMINISTRATO_GESTITO,
           QUANTITA,
           CAST(CASE WHEN SUBSTRING(SEGNO_OPERAZIONE, 1, 1) = 'U'
                     THEN -QUANTITA
                     ELSE  QUANTITA
                END AS DECIMAL(18, 5))                    AS QTA_CON_SEGNO
      FROM MOVIMENTI_TITOLI_EOR
     WHERE CONTO_AMMINISTRATO_GESTITO <> 'G-GESTITO'
       AND DATA_OPERAZIONE BETWEEN '1900/01/01' AND '2016/08/31'
) WITH DATA;

-- ---------------------------------------------------------------------
-- STEP 2 - Indice nell'ordine esatto delle window (PARTITION BY + ORDER BY),
--          con QTA_CON_SEGNO in coda: copre SUM/ROW_NUMBER senza leggere la tabella.
-- ---------------------------------------------------------------------
CREATE INDEX MOV_EOR_LAV_IX1
    ON MOV_EOR_LAV (CHIAVE_TITOLO,
                    CLIENTE_INTESTATARIO_XF,
                    DATA_OPERAZIONE,
                    NUMERO_OPERAZIONE,
                    QTA_CON_SEGNO);

-- ---------------------------------------------------------------------
-- STEP 3 - Query: nessun self-join. L'ultimo azzeramento si calcola con una
--          seconda window sulla stessa partizione invece di GROUP BY + JOIN.
-- ---------------------------------------------------------------------
WITH saldo AS (
    SELECT t.*,
           SUM(QTA_CON_SEGNO) OVER (
               PARTITION BY CHIAVE_TITOLO, CLIENTE_INTESTATARIO_XF
               ORDER BY DATA_OPERAZIONE, NUMERO_OPERAZIONE
               ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW
           ) AS SALDO_PROGRESSIVO,
           ROW_NUMBER() OVER (
               PARTITION BY CHIAVE_TITOLO, CLIENTE_INTESTATARIO_XF
               ORDER BY DATA_OPERAZIONE, NUMERO_OPERAZIONE
           ) AS PROG
      FROM MOV_EOR_LAV t
),
marcato AS (
    SELECT s.*,
           MAX(CASE WHEN SALDO_PROGRESSIVO = 0 THEN PROG END) OVER (
               PARTITION BY CHIAVE_TITOLO, CLIENTE_INTESTATARIO_XF
           ) AS PROG_ZERO,
           COUNT(*) OVER (
               PARTITION BY CHIAVE_TITOLO, CLIENTE_INTESTATARIO_XF
           ) AS N_MOVIMENTI
      FROM saldo s
)
SELECT NUMERO_OPERAZIONE,
       CODICE_ISIN,
       CODICE_INTERNO_TITOLO,
       SEGNO_OPERAZIONE,
       CONTO_AMMINISTRATO_GESTITO,
       CLIENTE_INTESTATARIO_XF,
       DATA_OPERAZIONE,
       QUANTITA,
       SALDO_PROGRESSIVO
  FROM marcato
 WHERE PROG <= PROG_ZERO            -- PROG_ZERO NULL (mai azzerato) -> escluso
-- AND PROG_ZERO = N_MOVIMENTI      -- VARIANTE: solo posizioni con saldo 0 AL 31/08/2016
 ORDER BY CHIAVE_TITOLO, CLIENTE_INTESTATARIO_XF, PROG;

-- ---------------------------------------------------------------------
-- STEP 4 (consigliato) - Materializzare il risultato invece di scaricarlo in ACS:
--   CREATE TABLE MOV_EOR_CICLI_CHIUSI AS ( <query dello STEP 3 senza ORDER BY> ) WITH DATA;
-- Misura il tempo reale del motore, senza il fetch JDBC di milioni di righe.
-- ---------------------------------------------------------------------
