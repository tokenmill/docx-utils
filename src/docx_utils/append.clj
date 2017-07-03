(ns docx-utils.append
  (:require [clojure.tools.logging :as log]
            [docx-utils.utils :refer [set-run]]
            [docx-utils.image :as image]
            [docx-utils.table :as table])
  (:import (org.apache.poi.xwpf.usermodel XWPFDocument XWPFParagraph XWPFPicture XWPFRun XWPFTable)
           (org.openxmlformats.schemas.wordprocessingml.x2006.main STTblWidth)))

(defn paragraph [^XWPFDocument document text]
  (log/debugf "Adding a paragraph '%s' to the end of the document." text)
  (-> document (.createParagraph) (.createRun)
      (set-run :text text)))

(defn ^XWPFPicture image [^XWPFDocument document image-path]
  (log/debugf "Adding an image '%s' to the end of the document." image-path)
  (-> document (.createParagraph) (.createRun)
      (image/insert image-path)))

(defn table [^XWPFDocument doc table-data]
  (log/debugf "Adding a table '%s' to the end of the document." table-data)
  (let [^XWPFTable table (.createTable doc)]
    (doto (-> table (.getCTTbl) (.addNewTblPr) (.addNewTblW))
      (.setType STTblWidth/DXA)
      (.setW (BigInteger/valueOf 9637)))
    (table/data-into-table table-data table)))

(defn bullet-list [^XWPFDocument doc list-data]
  (log/debugf "Adding a bullet list '%s' to the end of the document." list-data)
  ; TODO
  )

(defn numbered-list [^XWPFDocument doc list-data]
  (log/debugf "Adding a numbered list '%s' to the end of the document." list-data)
  ; TODO
  )