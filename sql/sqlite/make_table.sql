CREATE TABLE IF NOT EXISTS TBL_OHLCV (
  timestamp NUMERIC
  , updatetime NUMERIC
  , exchange TEXT
  , product_code TEXT
  , Open REAL
  , High REAL
  , Low REAL
  , Close REAL
  , Volume REAL
  , Change REAL
  , PRIMARY KEY (
    timestamp
    , exchange
    , product_code
  )
)