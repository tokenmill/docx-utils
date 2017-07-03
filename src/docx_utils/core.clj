(ns docx-utils.core
  (:require [clojure.string :as str]
            [clojure.java.io :as io]
            [clojure.tools.logging :as log])
  (:import (org.apache.poi.xwpf.extractor XWPFWordExtractor)
           (org.apache.poi POIXMLDocument)
           (org.apache.poi.xwpf.usermodel XWPFDocument XWPFParagraph XWPFRun XWPFTable XWPFTableRow XWPFPicture XWPFTableCell XWPFNumbering XWPFAbstractNum Document)
           (org.apache.poi.util Units)
           (javax.imageio ImageIO)
           (java.awt.image BufferedImage)
           (org.apache.xmlbeans XmlCursor)
           (org.openxmlformats.schemas.wordprocessingml.x2006.main CTTblWidth STTblWidth CTAbstractNum CTNumPr CTAbstractNum$Factory CTNum CTNum$Factory CTNumbering$Factory CTNumbering STHighlightColor STHighlightColor$Enum)
           (java.io ByteArrayOutputStream File)))

(def page-width-in-emu 6120000)


(defn clean-paragraph-content [^XWPFParagraph par]
  (doseq [run-index (range 0 (count (.getRuns par)))]
    (.removeRun par 0)))


(defn ^XWPFParagraph find-paragraph [^XWPFDocument doc ^String match]
  (->> doc
       (.getParagraphs)
       (filter (fn [^XWPFParagraph par] (= match (.getParagraphText par))))
       first))


(defn delete-paragraph [^XWPFDocument doc ^XWPFParagraph paragraph]
  (when-not (nil? paragraph)
    (.removeBodyElement doc (.getPosOfParagraph doc paragraph))))


(defn delete-placeholder-paragraph [^XWPFDocument doc ^String match]
  (when-let [par (find-paragraph doc match)]
    (delete-paragraph doc par)))


(defn set-run [^XWPFRun run & {:keys [text pos bold highlight-color]
                               :or {text ""
                                    pos 0
                                    highlight-color "none"}}]
  (.setText run (if (string? text) text (str text)) pos)
  (.setBold run (or bold false))
  (-> run (.getCTR) (.addNewRPr) (.addNewHighlight) (.setVal (STHighlightColor$Enum/forString highlight-color))))


(defn add-paragraph [^XWPFDocument doc text]
  (-> doc (.createParagraph) (.createRun)
      (set-run :text text)))


(defn replace-with-text-inside-paragraph
  "Text replacement based on XWPFRun class."
  [^XWPFDocument doc ^String match ^String replacement]
  (log/debugf "Replacing text '%s' with text '%s'" match replacement)
  (doseq [^XWPFParagraph par (.getParagraphs doc)]
    (doseq [^XWPFRun run (.getRuns par)]
      (when (and (.getText run 0)
                 (str/includes? (.getText run 0) match))
        (set-run run :text (str/replace (.getText run 0) match (str replacement)))))))


(defn replace-with-text
  "Text replacement based on XWPFParagraph class."
  [^XWPFDocument doc ^String match ^String replacement]
  (log/debugf "Replacing the paragraph '%s' with text '%s'" match replacement)
  (if (not (str/blank? replacement))
    (let [^XWPFParagraph par (find-paragraph doc match)]
      (clean-paragraph-content par)
      (set-run (.createRun par) :text replacement))
    (delete-placeholder-paragraph doc match)))


(defn image-size-in-emu [width-in-pixels height-in-pixels]
  (let [width-in-emu (min page-width-in-emu (Units/pixelToEMU width-in-pixels))
        height-in-emu (min page-width-in-emu (Units/pixelToEMU height-in-pixels))]
    (if (and (not= width-in-emu page-width-in-emu) (not= height-in-emu page-width-in-emu))
      [width-in-emu height-in-emu]
      (if (= width-in-emu page-width-in-emu)
        [width-in-emu (* height-in-pixels (/ page-width-in-emu  width-in-pixels))]
        [(* width-in-pixels (/ page-width-in-emu height-in-pixels)) height-in-emu]))))


(defn ^XWPFPicture put-image [^XWPFRun run image-path]
  (let [^BufferedImage bi (ImageIO/read (io/file image-path))
        [width-in-emu height-in-emu] (image-size-in-emu (.getWidth bi) (.getHeight bi))
        file-name (.getName (io/file image-path))]
    (with-open [image (io/input-stream image-path)]
      (.addPicture run
                   image
                   Document/PICTURE_TYPE_JPEG
                   file-name
                   width-in-emu
                   height-in-emu))))


(defn ^XWPFPicture add-image [^XWPFDocument doc image-path]
  (let [^XWPFParagraph paragraph (.createParagraph doc)
        ^XWPFRun run (.createRun paragraph)]
    (put-image run image-path)))


(defn replace-with-image [^XWPFDocument doc ^String match image-path]
  (log/debugf "Replacing the paragraph '%s' with image '%s'" match image-path)
  (let [^XWPFParagraph par (find-paragraph doc match)]
    (clean-paragraph-content par)
    (put-image (.createRun par) image-path)))


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


(defn add-table [^XWPFDocument doc table-data]
  (let [^XWPFTable table (.createTable doc)]
    (doto (-> table (.getCTTbl) (.addNewTblPr) (.addNewTblW))
      (.setType STTblWidth/DXA)
      (.setW (BigInteger/valueOf 9637)))
    (data-into-table table-data table)))


(defn replace-with-table
  "Given a placeholder string, inserts a table there."
  [^XWPFDocument doc ^String match table-data]
  (log/debugf "Replacing the paragraph '%s' with table '%s'" match table-data)
  (if (seq table-data)
    (let [^XWPFParagraph par (find-paragraph doc match)
          ^XWPFTable table (.insertNewTbl doc (.newCursor (.getCTP par)))]
      (doto (-> table (.getCTTbl) (.addNewTblPr) (.addNewTblW))
        (.setType STTblWidth/DXA)
        (.setW (BigInteger/valueOf 9637)))
      (data-into-table table-data table)
      (delete-paragraph doc par))
    (delete-placeholder-paragraph doc match)))


(def ^String cTAbstractNumBulletXML
  (str "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:abstractNumId=\"0\">"
       "<w:multiLevelType w:val=\"hybridMultilevel\"/>"
       "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"Symbol\" w:hAnsi=\"Symbol\" w:hint=\"default\"/></w:rPr></w:lvl>"
       "<w:lvl w:ilvl=\"1\" w:tentative=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"o\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"1440\" w:hanging=\"360\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"Courier New\" w:hAnsi=\"Courier New\" w:cs=\"Courier New\" w:hint=\"default\"/></w:rPr></w:lvl>"
       "<w:lvl w:ilvl=\"2\" w:tentative=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"2160\" w:hanging=\"360\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"Wingdings\" w:hAnsi=\"Wingdings\" w:hint=\"default\"/></w:rPr></w:lvl>"
       "</w:abstractNum>"))


(def ^String cTAbstractNumDecimalXML
  (str "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:abstractNumId=\"0\">"
       "<w:multiLevelType w:val=\"hybridMultilevel\"/>"
       "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/></w:pPr></w:lvl>"
       "<w:lvl w:ilvl=\"1\" w:tentative=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.%2\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"1440\" w:hanging=\"360\"/></w:pPr></w:lvl>"
       "<w:lvl w:ilvl=\"2\" w:tentative=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.%2.%3\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"2160\" w:hanging=\"360\"/></w:pPr></w:lvl>"
       "</w:abstractNum>"))


(defn replace-with-bullet-list [^XWPFDocument doc ^String match list-data]
  (log/debugf "Replacing the paragraph '%s' with list '%s'" match list-data)
  (if (seq list-data)
    (let [^XWPFParagraph placeholder-paragraph (find-paragraph doc match)
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
      (delete-paragraph doc placeholder-paragraph))
    (delete-placeholder-paragraph doc match)))


(defn ^XWPFDocument open-xwpf-doc [^String src]
  (XWPFDocument. (POIXMLDocument/openPackage src)))


(defn transform
  ([transformations]
   (let [template-file (.getPath (io/resource "template-1.docx"))]
     (transform template-file transformations)))
  ([template-file-path transformations]
   (when (nil? template-file-path) (throw (Exception. "Template file path is nil.")))
   (transform template-file-path (.getAbsolutePath (File/createTempFile "output-" ".docx")) transformations))
  ([template-file-path output-file-path transformations]
   (when (nil? template-file-path) (throw (Exception. "Template file path is nil.")))
   (when (nil? output-file-path) (throw (Exception. "Output file path is nil.")))
   (let [template (open-xwpf-doc template-file-path)]
     (log/infof "Applying transformations %s on template '%s' for output '%s'" transformations template-file-path output-file-path)

     (doseq [{:keys [type placeholder replacement]} transformations]
       (try
         (cond
           (= "text" type) (replace-with-text template placeholder replacement)
           (= "text_inline" type) (replace-with-text-inside-paragraph template placeholder replacement)
           (= "table" type) (replace-with-table template placeholder replacement)
           (= "image" type) (replace-with-image template placeholder replacement)
           (= "list" type) (replace-with-bullet-list template placeholder replacement)
           :else (log/warnf "Unknown transformation type '%s'" type))
         (catch Exception e
           (log/errorf "Failed transformation with type '%s' placeholder '%s' and replacement '%s'" type placeholder replacement))))

     (with-open [o (io/output-stream output-file-path)]
       (log/debugf "Writing the transformed template to the output file '%s'" output-file-path)
       (.write template o))
     output-file-path)))
