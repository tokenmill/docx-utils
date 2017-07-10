(ns docx-utils.elements.paragraph
  (:import (org.apache.poi.xwpf.usermodel XWPFParagraph XWPFDocument)))

(defn clean-paragraph-content [^XWPFParagraph par]
  (when-not (nil? par)
    (doseq [run-index (range 0 (count (.getRuns par)))]
      (.removeRun par 0))))

(defn ^XWPFParagraph find-paragraph [^XWPFDocument doc ^String match]
  (let [top-level-paragraphs (-> doc (.getParagraphs))
        paragraphs-in-tables (->> doc
                                  (.getTables)
                                  (map #(.getRows %))
                                  (reduce #(into %1 %2) [])
                                  (map #(.getTableICells %))
                                  (reduce #(into %1 %2) [])
                                  (map #(.getParagraphs %))
                                  (reduce #(into %1 %2) []))]
    (->> top-level-paragraphs
         (concat paragraphs-in-tables)
         (filter (fn [^XWPFParagraph par] (= match (.getParagraphText par))))
         first)))

(defn delete-paragraph [^XWPFDocument doc ^XWPFParagraph paragraph]
  (when-not (nil? paragraph)
    (.removeBodyElement doc (.getPosOfParagraph doc paragraph))))

(defn delete-placeholder-paragraph [^XWPFDocument doc ^String match]
  (when-let [par (find-paragraph doc match)]
    (delete-paragraph doc par)))