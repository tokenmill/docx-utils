(ns docx-utils.elements.listing
  (:require [docx-utils.constants :refer [cTAbstractNumBulletXML cTAbstractNumDecimalXML]]
            [docx-utils.elements.paragraph :as paragraph]
            [docx-utils.elements.run :refer [set-run]])
  (:import (org.apache.poi.xwpf.usermodel XWPFDocument XWPFParagraph XWPFAbstractNum XWPFNumbering)
           (org.openxmlformats.schemas.wordprocessingml.x2006.main CTNumbering CTAbstractNum CTNumbering$Factory)))

(defn bullet-list [^XWPFDocument doc ^XWPFParagraph paragraph list-data]
  (let [^CTNumbering cTNumbering (CTNumbering$Factory/parse cTAbstractNumBulletXML)
        ^CTAbstractNum cTAbstractNum (.getAbstractNumArray cTNumbering 0)
        ^XWPFAbstractNum abstractNum (XWPFAbstractNum. cTAbstractNum)
        ^XWPFNumbering numbering (.createNumbering doc)
        ^BigInteger abstractNumID (.addAbstractNum numbering abstractNum)
        ^BigInteger numID (.addNum numbering abstractNumID)
        highlight-colors (-> list-data (meta) :highlight-colors)
        bolds (-> list-data (meta) :bolds)]
    (doseq [[index item] (map vector (iterate inc 0) list-data)]
      (let [^XWPFParagraph item-paragraph (.insertNewParagraph doc (.newCursor (.getCTP paragraph)))]
        (.setNumID item-paragraph numID)
        (.setStyle item-paragraph "ListParagraph")
        (-> item-paragraph (.getCTP) (.getPPr) (.getNumPr) (.addNewIlvl) (.setVal (BigInteger/ZERO)))
        (set-run (.createRun item-paragraph) :text item :bold (nth bolds index) :highlight-color (nth highlight-colors index))))
    (paragraph/delete-paragraph doc paragraph)))

(defn numbered-list [^XWPFDocument doc ^XWPFParagraph paragraph list-data]
  (let [^CTNumbering cTNumbering (CTNumbering$Factory/parse cTAbstractNumDecimalXML)
        ^CTAbstractNum cTAbstractNum (.getAbstractNumArray cTNumbering 0)
        ^XWPFAbstractNum abstractNum (XWPFAbstractNum. cTAbstractNum)
        ^XWPFNumbering numbering (.createNumbering doc)
        ^BigInteger abstractNumID (.addAbstractNum numbering abstractNum)
        ^BigInteger numID (.addNum numbering abstractNumID)
        highlight-colors (-> list-data (meta) :highlight-colors)
        bolds (-> list-data (meta) :bolds)]
    (doseq [[index item] (map vector (iterate inc 0) list-data)]
      (let [^XWPFParagraph item-paragraph (.insertNewParagraph doc (.newCursor (.getCTP paragraph)))]
        (.setNumID item-paragraph numID)
        (.setStyle item-paragraph "ListParagraph")
        (-> item-paragraph (.getCTP) (.getPPr) (.getNumPr) (.addNewIlvl) (.setVal (BigInteger/ZERO)))
        (set-run (.createRun item-paragraph) :text item :bold (nth bolds index) :highlight-color (nth highlight-colors index))))
    (paragraph/delete-paragraph doc paragraph)))
