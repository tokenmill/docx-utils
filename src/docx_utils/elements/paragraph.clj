(ns docx-utils.elements.paragraph
  (:import (org.apache.poi.xwpf.usermodel XWPFParagraph XWPFDocument)))

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