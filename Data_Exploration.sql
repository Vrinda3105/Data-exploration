-- Basic Statistics
SELECT
    MIN([Date]) AS min_date,
    MAX([Date]) AS max_date,
    AVG([Close]) AS avg_close_price,
    STDEV([Close]) AS price_volatility,
    MAX(High) AS max_high_price,
    MIN(Low) AS min_low_price
FROM SQLproject.dbo.Sheet1$;

-- Top Daily Gainers
SELECT
    [Date],
    [Close],
    LAG([Close]) OVER (ORDER BY [Date]) AS previous_day_close,
    [Close] - LAG([Close]) OVER (ORDER BY [Date]) AS daily_change
FROM SQLproject.dbo.Sheet1$
ORDER BY daily_change DESC
OFFSET 0 ROWS
FETCH NEXT 5 ROWS ONLY;

-- Daily Trading Volume Trends
SELECT
    [Date],
    Volume,
    AVG(Volume) OVER (ORDER BY [Date] ROWS BETWEEN 29 PRECEDING AND CURRENT ROW) AS avg_30day_volume
FROM SQLproject.dbo.Sheet1$
ORDER BY [Date];

-- Moving Averages
SELECT
    [Date],
    [Close],
    AVG([Close]) OVER (ORDER BY [Date] ROWS BETWEEN 6 PRECEDING AND CURRENT ROW) AS [7day_moving_avg]
FROM SQLproject.dbo.Sheet1$
ORDER BY [Date];

-- Exponential Moving Average (EMA)
SELECT
    [Date],
    [Close],
    AVG([Close]) OVER (ORDER BY [Date] ROWS BETWEEN 29 PRECEDING AND CURRENT ROW) AS [30day_simple_moving_avg],
    EXP(SUM(LOG([Close] + 1)) OVER (ORDER BY [Date] ROWS BETWEEN 29 PRECEDING AND CURRENT ROW) / 30) - 1 AS [30day_ema]
FROM SQLproject.dbo.Sheet1$
ORDER BY [Date];

-- Identifying Price Trends
SELECT
    [Date],
    [Close],
    CASE WHEN [Close] > AVG([Close]) OVER (ORDER BY [Date] ROWS BETWEEN 29 PRECEDING AND CURRENT ROW) THEN 'Upward'
         WHEN [Close] < AVG([Close]) OVER (ORDER BY [Date] ROWS BETWEEN 29 PRECEDING AND CURRENT ROW) THEN 'Downward'
         ELSE 'Neutral' END AS price_trend
FROM SQLproject.dbo.Sheet1$
ORDER BY [Date];

-- RSI Signals
WITH RSI_CTE AS (
    SELECT
        [Date],
        [Close],
        AVG(Gain) OVER (ORDER BY [Date] ROWS BETWEEN 14 PRECEDING AND CURRENT ROW) AS avg_gain,
        AVG(Loss) OVER (ORDER BY [Date] ROWS BETWEEN 14 PRECEDING AND CURRENT ROW) AS avg_loss
    FROM (
        SELECT
            [Date],
            [Close],
            CASE WHEN [Close] > LAG([Close]) OVER (ORDER BY [Date]) THEN [Close] - LAG([Close]) OVER (ORDER BY [Date]) ELSE 0 END AS Gain,
            CASE WHEN [Close] < LAG([Close]) OVER (ORDER BY [Date]) THEN LAG([Close]) OVER (ORDER BY [Date]) - [Close] ELSE 0 END AS Loss
        FROM SQLproject.dbo.Sheet1$
    ) AS RSI_Calculation
)
SELECT
    [Date],
    [Close],
    100 - (100 / (1 + (avg_gain / NULLIF(avg_loss, 0)))) AS rsi,
    CASE WHEN (100 - (100 / (1 + (avg_gain / NULLIF(avg_loss, 0))))) > 70 THEN 'Overbought'
         WHEN (100 - (100 / (1 + (avg_gain / NULLIF(avg_loss, 0))))) < 30 THEN 'Oversold'
         ELSE 'Neutral' END AS rsi_signal
FROM RSI_CTE
ORDER BY [Date];

-- MACD Signals
WITH MACD_CTE AS (
    SELECT
        [Date],
        [Close],
        AVG([Close]) OVER (ORDER BY [Date] ROWS BETWEEN 11 PRECEDING AND CURRENT ROW) AS twelve_day_ema,
        AVG([Close]) OVER (ORDER BY [Date] ROWS BETWEEN 25 PRECEDING AND CURRENT ROW) AS twenty_six_day_ema
    FROM SQLproject.dbo.Sheet1$
)
SELECT
    [Date],
    [Close],
    twelve_day_ema,
    twenty_six_day_ema,
    twelve_day_ema - twenty_six_day_ema AS macd,
    AVG(twelve_day_ema - twenty_six_day_ema) OVER (ORDER BY [Date] ROWS BETWEEN 8 PRECEDING AND CURRENT ROW) AS signal_line
FROM MACD_CTE
ORDER BY [Date];