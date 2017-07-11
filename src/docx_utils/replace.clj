(ns docx-utils.replace
  (:require [clojure.tools.logging :as log]
            [clojure.string :as str]
            [docx-utils.elements.paragraph :as paragraph]
            [docx-utils.elements.run :refer [set-run]]
            [docx-utils.elements.image :as image]
            [docx-utils.elements.table :as table]
            [docx-utils.elements.listing :as listing])
  (:import (org.apache.poi.xwpf.usermodel XWPFDocument XWPFParagraph XWPFRun XWPFAbstractNum XWPFNumbering XWPFTable)
           (org.openxmlformats.schemas.wordprocessingml.x2006.main CTNumbering CTAbstractNum CTNumbering$Factory STTblWidth)))

(defn with-text-inline
  "Text replacement based on XWPFRun class."
  [^XWPFDocument doc ^String match ^String replacement]
  (log/debugf "Replacing text '%s' with text '%s'" match replacement)
  (doseq [^XWPFParagraph par (paragraph/paragraphs doc)]
    (doseq [^XWPFRun run (.getRuns par)]
      (when (and (.getText run 0)
                 (str/includes? (.getText run 0) match))
        (set-run run :text (str/replace (.getText run 0) match (str replacement)))))))

(defn with-text
  "Text replacement based on XWPFParagraph class."
  [^XWPFDocument doc ^String match ^String replacement]
  (log/debugf "Replacing the paragraph '%s' with text '%s'" match replacement)
  (if (not (str/blank? replacement))
    (when-let [^XWPFParagraph par (paragraph/find-paragraph doc match)]
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
    (let [^XWPFParagraph placeholder-paragraph (paragraph/find-paragraph doc match)]
      (listing/bullet-list doc placeholder-paragraph list-data))
    (paragraph/delete-placeholder-paragraph doc match)))

(defn with-numbered-list [^XWPFDocument doc ^String match list-data]
  (log/debugf "Replacing the paragraph '%s' with a numbered list '%s'" match list-data)
  (if (seq list-data)
    (let [^XWPFParagraph placeholder-paragraph (paragraph/find-paragraph doc match)]
      (listing/numbered-list doc placeholder-paragraph list-data))
    (paragraph/delete-placeholder-paragraph doc match)))

(defn with-table
  "Given a placeholder string, inserts a table there."
  [^XWPFDocument doc ^String match table-data]
  (log/debugf "Replacing the paragraph '%s' with table '%s'" match (pr-str table-data))
  (if (seq table-data)
    (let [^XWPFParagraph par (paragraph/find-paragraph doc match)
          ^XWPFTable table (.insertNewTbl doc (.newCursor (.getCTP par)))]
      (table/fix-width table)
      (table/data-into-table table-data table)
      (paragraph/delete-paragraph doc par))
    (paragraph/delete-placeholder-paragraph doc match)))
