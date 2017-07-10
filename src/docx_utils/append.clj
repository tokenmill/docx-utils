(ns docx-utils.append
  (:require [clojure.tools.logging :as log]
            [docx-utils.elements.run :refer [set-run]]
            [docx-utils.elements.listing :as listing]
            [docx-utils.elements.image :as image]
            [docx-utils.elements.table :as table])
  (:import (org.apache.poi.xwpf.usermodel XWPFDocument XWPFParagraph XWPFPicture XWPFRun XWPFTable)
           (org.openxmlformats.schemas.wordprocessingml.x2006.main STTblWidth)))

(defn text [^XWPFDocument document value]
  (log/debugf "Adding a text '%s' to the end of the document." value)
  (-> document (.createParagraph) (.createRun)
      (set-run :text value)))

(defn text-inline [^XWPFDocument document value]
  (log/debugf "Adding an inline text '%s' to the end of the document." value)
  (if-let [last-paragraph (some-> document
                              (.getParagraphs)
                              (last))]
    (-> last-paragraph (.createRun) (set-run :text value))
    (text document value)))

(defn ^XWPFPicture image [^XWPFDocument document image-path]
  (log/debugf "Adding an image '%s' to the end of the document." image-path)
  (-> document (.createParagraph) (.createRun)
      (image/insert image-path)))

(defn table [^XWPFDocument document table-data]
  (log/debugf "Adding a table '%s' to the end of the document." table-data)
  (let [^XWPFTable table (.createTable document)]
    (table/fix-width table)
    (table/data-into-table table-data table)))

(defn bullet-list [^XWPFDocument document list-data]
  (log/debugf "Adding a bullet list '%s' to the end of the document." (pr-str list-data))
  (let [paragraph (-> document (.createParagraph))]
    (.createRun paragraph)
    (listing/bullet-list document paragraph list-data)))

(defn numbered-list [^XWPFDocument document list-data]
  (log/debugf "Adding a numbered list '%s' to the end of the document." (pr-str list-data))
  (let [paragraph (-> document (.createParagraph))]
    (.createRun paragraph)
    (listing/numbered-list document paragraph list-data)))
