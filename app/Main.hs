{-# LANGUAGE OverloadedStrings #-}
module Main (main) where

import Codec.Xlsx
import Control.Lens
import qualified Data.ByteString.Lazy as L
import Data.Time.Clock.POSIX

main :: IO ()
main = do
  -- Read template 1
  t1 <- toXlsx <$> L.readFile "template1.xlsx"
  -- Read cell B3 from Sheet1
  let value = t1 ^? ixSheet "Sheet1" . ixCell (3,2) . cellValue . _Just
  print value
  -- Read template 2
  t2 <- toXlsx <$> L.readFile "template2.xlsx"
  -- Copy value read from t1 to t2, Sheet1, cell C4
  case value of
    Nothing -> putStrLn "No value?"
    Just v  -> do
      case (t2 ^? ixSheet "Sheet1") of
        Nothing -> putStrLn "No Sheet1 in template 2?"
        Just s  -> do
          let s' = copyTo s (4,3) v
          -- Write modified file back to output.xlsx
          ct <- getPOSIXTime
          let t2' = t2 & atSheet "Sheet1" ?~ s'
          L.writeFile "output.xlsx" (fromXlsx ct t2')

  putStrLn "Finished."

copyTo :: Worksheet -> (Int, Int) -> CellValue -> Worksheet
copyTo sheet posn value = sheet & cellValueAt posn ?~ value