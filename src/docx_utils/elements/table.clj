(ns docx-utils.elements.table
  (:require [clojure.tools.logging :as log]
            [docx-utils.elements.run :refer [set-run]])
  (:import (org.apache.poi.xwpf.usermodel XWPFTable XWPFTableRow XWPFTableCell XWPFRun)
           (org.openxmlformats.schemas.wordprocessingml.x2006.main STTblWidth)))

(defn fix-width [^XWPFTable table]
  (doto (-> table (.getCTTbl) (.addNewTblPr) (.addNewTblW))
    (.setType STTblWidth/DXA)
    (.setW (BigInteger/valueOf 9637))))

(defn data-into-table [table-data ^XWPFTable table]
  (log/debugf "Adding data to the table")
  (doseq [[^Integer line-index line] (map vector (iterate inc 0) table-data)]
    (let [^XWPFTableRow row (or (.getRow table line-index) (.createRow table))
          color (or (-> table-data (nth line-index) (meta) :color) "FFFFFF")
          cell-boldness (or (-> table-data (nth line-index) (meta) :bold) false)]
      (doseq [[cell-index cell-value] (map vector (iterate inc 0) line)]
        (let [^XWPFTableCell cell (or (.getCell row cell-index) (.createCell row))
              ^XWPFRun run (-> cell (.getParagraphs) (first) (.createRun))]
          (set-run run :text cell-value :bold cell-boldness)
          (.setColor cell color))))))
