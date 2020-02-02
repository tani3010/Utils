INSERT OR IGNORE INTO TBL_OHLCV (
  timestamp
  , updatetime
  , exchange
  , product_code
  , Open
  , High
  , Low
  , Close
  , Volume
  , Change
)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)