(ns docx-utils.replace
  (:require [docx-utils.paragraph :as paragraph]
            [clojure.tools.logging :as log]
            [clojure.string :as str]
            [docx-utils.constants :refer [cTAbstractNumBulletXML]]
            [docx-utils.utils :refer [set-run]]
            [docx-utils.image :as image]
            [docx-utils.table :as table])
  (:import (org.apache.poi.xwpf.usermodel XWPFDocument XWPFParagraph XWPFRun XWPFAbstractNum XWPFNumbering XWPFTable)
           (org.openxmlformats.schemas.wordprocessingml.x2006.main CTNumbering CTAbstractNum CTNumbering$Factory STTblWidth)))

(defn with-inline-text
  "Text replacement based on XWPFRun class."
  [^XWPFDocument doc ^String match ^String replacement]
  (log/debugf "Replacing text '%s' with text '%s'" match replacement)
  (doseq [^XWPFParagraph par (.getParagraphs doc)]
    (doseq [^XWPFRun run (.getRuns par)]
      (when (and (.getText run 0)
                 (str/includes? (.getText run 0) match))
        (set-run run :text (str/replace (.getText run 0) match (str replacement)))))))

(defn with-text
  "Text replacement based on XWPFParagraph class."
  [^XWPFDocument doc ^String match ^String replacement]
  (log/debugf "Replacing the paragraph '%s' with text '%s'" match replacement)
  (if (not (str/blank? replacement))
    (let [^XWPFParagraph par (paragraph/find-paragraph doc match)]
      (paragraph/clean-paragraph-content par)
      (set-run (.createRun par) :text replacement))
    (paragraph/delete-placeholder-paragraph doc match)))


(defn with-image [^XWPFDocument doc ^String match image-path]
  (log/debugf "Replacing the paragraph '%s' with image '%s'" match image-path)
  (let [^XWPFParagraph par (paragraph/find-paragraph doc match)]
    (paragraph/clean-paragraph-content par)
    (image/insert (.createRun par) image-path)))

(defn with-bullet-list [^XWPFDocument doc ^String match list-data]
  (log/debugf "Replacing the paragraph '%s' with list '%s'" match list-data)
  (if (seq list-data)
    (let [^XWPFParagraph placeholder-paragraph (paragraph/find-paragraph doc match)
          ^CTNumbering cTNumbering (CTNumbering$Factory/parse cTAbstractNumBulletXML)
          ^CTAbstractNum cTAbstractNum (.getAbstractNumArray cTNumbering 0)
          ^XWPFAbstractNum abstractNum (XWPFAbstractNum. cTAbstractNum)
          ^XWPFNumbering numbering (.createNumbering doc)
          ^BigInteger abstractNumID (.addAbstractNum numbering abstractNum)
          ^BigInteger numID (.addNum numbering abstractNumID)
          highlight-colors (-> list-data (meta) :highlight-colors)
          bolds (-> list-data (meta) :bolds)]
      (doseq [[index item] (map vector (iterate inc 0) list-data)]
        (let [^XWPFParagraph item-paragraph (.insertNewParagraph doc (.newCursor (.getCTP placeholder-paragraph)))]
          (.setNumID item-paragraph numID)
          (.setStyle item-paragraph "ListParagraph")
          (-> item-paragraph (.getCTP) (.getPPr) (.getNumPr) (.addNewIlvl) (.setVal (BigInteger/ZERO)))
          (set-run (.createRun item-paragraph) :text item :bold (nth bolds index) :highlight-color (nth highlight-colors index))))
      (paragraph/delete-paragraph doc placeholder-paragraph))
    (paragraph/delete-placeholder-paragraph doc match)))

(defn with-table
  "Given a placeholder string, inserts a table there."
  [^XWPFDocument doc ^String match table-data]
  (log/debugf "Replacing the paragraph '%s' with table '%s'" match table-data)
  (if (seq table-data)
    (let [^XWPFParagraph par (paragraph/find-paragraph doc match)
          ^XWPFTable table (.insertNewTbl doc (.newCursor (.getCTP par)))]
      (doto (-> table (.getCTTbl) (.addNewTblPr) (.addNewTblW))
        (.setType STTblWidth/DXA)
        (.setW (BigInteger/valueOf 9637)))
      (table/data-into-table table-data table)
      (paragraph/delete-paragraph doc par))
    (paragraph/delete-placeholder-paragraph doc match)))
