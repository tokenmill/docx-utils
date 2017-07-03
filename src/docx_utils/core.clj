(ns docx-utils.core
  (:require [clojure.tools.logging :as log]
            [docx-utils.io :as docx-io]
            [docx-utils.append :as append]
            [docx-utils.replace :as replace])
  (:import (org.apache.poi.xwpf.usermodel XWPFDocument)))

(defn apply-transformation [^XWPFDocument document {:keys [type placeholder replacement]}]
  (log/infof "Applying transformation of type '%s' for placeholder '%s' with replacement '%s'" type placeholder replacement)
  (try
    (cond
      (= :append-text type) (append/paragraph document replacement)
      (= :append-image type) (append/image document replacement)
      (= :append-table type) (append/table document replacement)
      (= :append-bullet-list type) (append/bullet-list document replacement)
      (= :append-numbered-list type) (append/numbered-list document replacement)
      (= "text" type) (replace/with-text document placeholder replacement)
      (= "text_inline" type) (replace/with-inline-text document placeholder replacement)
      (= "table" type) (replace/with-table document placeholder replacement)
      (= "image" type) (replace/with-image document placeholder replacement)
      (= "list" type) (replace/with-bullet-list document placeholder replacement)
      :else (log/warnf "Unknown transformation type '%s'" type))
    (catch Exception e
      (log/errorf "Failed transformation with type '%s' placeholder '%s' and replacement '%s'" type placeholder replacement))))

(defn apply-transformations [^XWPFDocument document transformations]
  (doseq [transformation transformations]
    (apply-transformation document transformation)))

(defn transform
  ([transformations]
   (transform nil transformations))
  ([template-file-path transformations]
   (transform template-file-path (docx-utils.io/temp-output-file) transformations))
  ([template-file-path output-file-path transformations]
   (when (nil? output-file-path) (throw (Exception. "Output file path is nil.")))
   (let [^XWPFDocument document (docx-io/load-template template-file-path)]
     (log/infof "Applying transformations %s on template '%s' for output '%s'"
                transformations template-file-path output-file-path)
     (apply-transformations document transformations)
     (docx-io/store document output-file-path)
     output-file-path)))
