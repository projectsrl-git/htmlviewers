-- Operazioni per titolo + CLIENTE_INTESTATARIO_XF che, partendo da saldo 0,
-- riportano il saldo progressivo a 0 entro il 31/08/2016.
-- Restituisce tutte le operazioni fino all'ULTIMO azzeramento avvenuto entro la data.
--
-- Db2 for i 7.4. MOVIMENTI_TITOLI_EOR e' una CTE: un solo WITH per statement,
-- le CTE seguenti vanno concatenate dopo la sua ")" sostituendo la SELECT finale.
--
-- Controllo segni (attesi solo 'E - ENTRATA' / 'U - USCITA'):
--   SELECT SEGNO_OPERAZIONE, COUNT(*) FROM MOVIMENTI_TITOLI_EOR GROUP BY SEGNO_OPERAZIONE;

WITH MOVIMENTI_TITOLI_EOR AS (
    /* ... definizione esistente ... */
),
mov AS (
    SELECT m.NUMERO_OPERAZIONE,
           m.CODICE_INTERNO_TITOLO,
           m.CODICE_ISIN,
           -- ISIN vuoto (ELSE ' ' nella CTE) -> fallback sul codice interno,
           -- altrimenti tutti i titoli senza ISIN finirebbero nello stesso saldo
           COALESCE(NULLIF(TRIM(m.CODICE_ISIN), ''), 'INT:' CONCAT TRIM(m.CODICE_INTERNO_TITOLO)) AS CHIAVE_TITOLO,
           m.SEGNO_OPERAZIONE,
           m.CONTO_AMMINISTRATO_GESTITO,
           m.CLIENTE_INTESTATARIO_XF,
           m.DATA_OPERAZIONE,
           m.QUANTITA,
           -- SEGNO_OPERAZIONE: 'E - ENTRATA' / 'U - USCITA'
           CASE WHEN SUBSTRING(m.SEGNO_OPERAZIONE, 1, 1) = 'U'
                THEN -m.QUANTITA
                ELSE  m.QUANTITA
           END AS QTA_CON_SEGNO
      FROM MOVIMENTI_TITOLI_EOR m
     WHERE m.CONTO_AMMINISTRATO_GESTITO <> 'G-GESTITO'
       -- DATA_OPERAZIONE e' VARCHAR(10) 'AAAA/MM/GG': confronto lessicografico = cronologico,
       -- nessuna conversione a DATE (che causava SQL0181). Il limite basso scarta blank/zeri.
       AND m.DATA_OPERAZIONE BETWEEN '1900/01/01' AND '2016/08/31'
),
saldo AS (
    SELECT mov.*,
           SUM(QTA_CON_SEGNO) OVER (
               PARTITION BY CHIAVE_TITOLO, CLIENTE_INTESTATARIO_XF
               ORDER BY DATA_OPERAZIONE, NUMERO_OPERAZIONE
               ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW
           ) AS SALDO_PROGRESSIVO,
           ROW_NUMBER() OVER (
               PARTITION BY CHIAVE_TITOLO, CLIENTE_INTESTATARIO_XF
               ORDER BY DATA_OPERAZIONE, NUMERO_OPERAZIONE
           ) AS PROG
      FROM mov
),
ultimo_zero AS (
    SELECT CHIAVE_TITOLO,
           CLIENTE_INTESTATARIO_XF,
           MAX(PROG) AS PROG_ZERO
      FROM saldo
     WHERE SALDO_PROGRESSIVO = 0
     GROUP BY CHIAVE_TITOLO, CLIENTE_INTESTATARIO_XF
)
SELECT s.NUMERO_OPERAZIONE,
       s.CODICE_ISIN,
       s.CODICE_INTERNO_TITOLO,
       s.SEGNO_OPERAZIONE,
       s.CONTO_AMMINISTRATO_GESTITO,
       s.CLIENTE_INTESTATARIO_XF,
       s.DATA_OPERAZIONE,
       s.QUANTITA,
       s.SALDO_PROGRESSIVO
  FROM saldo s
  JOIN ultimo_zero z
    ON z.CHIAVE_TITOLO           = s.CHIAVE_TITOLO
   AND z.CLIENTE_INTESTATARIO_XF = s.CLIENTE_INTESTATARIO_XF
   AND s.PROG                   <= z.PROG_ZERO
 ORDER BY s.CHIAVE_TITOLO, s.CLIENTE_INTESTATARIO_XF, s.PROG;

-- VARIANTE "posizione chiusa al 31/08/2016" (saldo alla data esattamente 0):
-- sostituire ultimo_zero con
--
-- ultimo_zero AS (
--     SELECT CHIAVE_TITOLO, CLIENTE_INTESTATARIO_XF, MAX(PROG) AS PROG_ZERO
--       FROM saldo
--      GROUP BY CHIAVE_TITOLO, CLIENTE_INTESTATARIO_XF
--     HAVING MAX(CASE WHEN SALDO_PROGRESSIVO = 0 THEN PROG END) = MAX(PROG)
-- )
