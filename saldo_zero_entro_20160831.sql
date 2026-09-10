-- Operazioni per CODICE_ISIN + CLIENTE_INTESTATARIO_XF che, partendo da saldo 0,
-- riportano il saldo progressivo a 0 entro il 31/08/2016.
-- Restituisce tutte le operazioni fino all'ULTIMO azzeramento avvenuto entro la data.
--
-- NB: MOVIMENTI_TITOLI_EOR e' una CTE. Un solo WITH per statement:
--     incollare questo blocco SUBITO DOPO la ")" che chiude MOVIMENTI_TITOLI_EOR,
--     sostituendo la SELECT finale esistente.

WITH MOVIMENTI_TITOLI_EOR AS (
    /* ... definizione esistente ... */
),
mov AS (
    SELECT m.NUMERO_OPERAZIONE,
           m.CODICE_ISIN,
           m.SEGNO_OPERAZIONE,
           m.CONTO_AMMINISTRATO_GESTITO,
           m.CLIENTE_INTESTATARIO_XF,
           m.DATA_OPERAZIONE,
           m.QUANTITA,
           CASE WHEN m.SEGNO_OPERAZIONE = '-'          -- <-- ADATTARE al valore reale di "scarico"
                THEN -m.QUANTITA
                ELSE  m.QUANTITA
           END AS QTA_CON_SEGNO
      FROM MOVIMENTI_TITOLI_EOR m
     WHERE m.CONTO_AMMINISTRATO_GESTITO <> 'G-GESTITO'
       AND m.DATA_OPERAZIONE < DATE '2016-09-01'       -- SQL Server: < '2016-09-01'
),
saldo AS (
    SELECT mov.*,
           SUM(QTA_CON_SEGNO) OVER (
               PARTITION BY CODICE_ISIN, CLIENTE_INTESTATARIO_XF
               ORDER BY DATA_OPERAZIONE, NUMERO_OPERAZIONE
               ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW
           ) AS SALDO_PROGRESSIVO,
           ROW_NUMBER() OVER (
               PARTITION BY CODICE_ISIN, CLIENTE_INTESTATARIO_XF
               ORDER BY DATA_OPERAZIONE, NUMERO_OPERAZIONE
           ) AS PROG
      FROM mov
),
ultimo_zero AS (
    SELECT CODICE_ISIN,
           CLIENTE_INTESTATARIO_XF,
           MAX(PROG) AS PROG_ZERO
      FROM saldo
     WHERE SALDO_PROGRESSIVO = 0
     GROUP BY CODICE_ISIN, CLIENTE_INTESTATARIO_XF
)
SELECT s.NUMERO_OPERAZIONE,
       s.CODICE_ISIN,
       s.SEGNO_OPERAZIONE,
       s.CONTO_AMMINISTRATO_GESTITO,
       s.CLIENTE_INTESTATARIO_XF,
       s.DATA_OPERAZIONE,
       s.QUANTITA,
       s.SALDO_PROGRESSIVO
  FROM saldo s
  JOIN ultimo_zero z
    ON z.CODICE_ISIN             = s.CODICE_ISIN
   AND z.CLIENTE_INTESTATARIO_XF = s.CLIENTE_INTESTATARIO_XF
   AND s.PROG                   <= z.PROG_ZERO
 ORDER BY s.CODICE_ISIN, s.CLIENTE_INTESTATARIO_XF, s.PROG;

-- VARIANTE "posizione chiusa al 31/08/2016":
-- per tenere solo le coppie il cui saldo alla data è esattamente 0
-- (nessuna operazione successiva all'ultimo azzeramento), sostituire ultimo_zero con:
--
-- ultimo_zero AS (
--     SELECT CODICE_ISIN, CLIENTE_INTESTATARIO_XF, MAX(PROG) AS PROG_ZERO
--       FROM saldo
--      GROUP BY CODICE_ISIN, CLIENTE_INTESTATARIO_XF
--     HAVING MAX(CASE WHEN SALDO_PROGRESSIVO = 0 THEN PROG END) = MAX(PROG)
-- )
