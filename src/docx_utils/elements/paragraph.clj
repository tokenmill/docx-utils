(ns docx-utils.elements.paragraph
  (:import (org.apache.poi.xwpf.usermodel XWPFParagraph XWPFDocument)))

(defn paragraphs
  "Given a  document extracts paragraphs from various places and concatenates
  them into one list."
  [^XWPFDocument doc]
  (concat
    (->> doc (.getParagraphs))
    (->> doc
         (.getTables)
         (map #(.getRows %))
         (reduce #(into %1 %2) [])
         (map #(.getTableICells %))
         (reduce #(into %1 %2) [])
         (map #(.getParagraphs %))
         (reduce #(into %1 %2) []))))

(defn clean-paragraph-content [^XWPFParagraph par]
  (when-not (nil? par)
    (doseq [run-index (range 0 (count (.getRuns par)))]
      (.removeRun par 0))))

(defn ^XWPFParagraph find-paragraph [^XWPFDocument doc ^String match]
  (->> doc
       (paragraphs)
       (filter (fn [^XWPFParagraph par] (= match (.getParagraphText par))))
       first))

(defn delete-paragraph [^XWPFDocument doc ^XWPFParagraph paragraph]
  (when-not (nil? paragraph)
    (.removeBodyElement doc (.getPosOfParagraph doc paragraph))))

(defn delete-placeholder-paragraph [^XWPFDocument doc ^String match]
  (when-let [par (find-paragraph doc match)]
    (delete-paragraph doc par)))