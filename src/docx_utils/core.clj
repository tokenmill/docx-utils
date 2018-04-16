(ns docx-utils.core
  (:require [clojure.java.io :as io]
            [clojure.tools.logging :as log]
            [docx-utils.io :as docx-io]
            [docx-utils.append :as append]
            [docx-utils.replace :as replace])
  (:import (org.apache.poi.xwpf.usermodel XWPFDocument)))

(defn decorate-placeholder [undecorated-placeholder]
  (str "%{" undecorated-placeholder "}"))

(defn apply-transformation [^XWPFDocument document {:keys [type placeholder replacement]}]
  (let [decorated-placeholder (decorate-placeholder placeholder)]
    (log/infof "Applying transformation of type '%s' for placeholder '%s' with replacement '%s'"
               type decorated-placeholder replacement)
    (try
      (cond
        (= :append-text type) (append/text document replacement)
        (= :append-text-inline type) (append/text-inline document replacement)
        (= :append-image type) (append/image document replacement)
        (= :append-table type) (append/table document replacement)
        (= :append-bullet-list type) (append/bullet-list document replacement)
        (= :append-numbered-list type) (append/numbered-list document replacement)
        (= :replace-text type) (replace/with-text document decorated-placeholder replacement)
        (= :replace-text-inline type) (replace/with-text-inline document decorated-placeholder replacement)
        (= :replace-table type) (replace/with-table document decorated-placeholder replacement)
        (= :replace-image type) (replace/with-image document decorated-placeholder replacement)
        (= :replace-bullet-list type) (replace/with-bullet-list document decorated-placeholder replacement)
        (= :replace-numbered-list type) (replace/with-numbered-list document decorated-placeholder replacement)
        :else (log/warnf "Unknown transformation type '%s'" type))
      (catch Exception e
        (log/errorf "Failed transformation with type '%s' decorated-placeholder '%s' and replacement '%s' with exception: %s"
                    type decorated-placeholder replacement (.printStackTrace e))))))

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

(defn transform-input-stream
  "Retrieves document from given InputStream.
  Transforms the stream data and returns a new InputStream."
  ([transformations ^java.io.InputStream doc-input-stream]
   (with-open [^XWPFDocument document (docx-io/load-template-from-memory doc-input-stream)]
     (apply-transformations document transformations)
     (with-open [output-stream (java.io.ByteArrayOutputStream.)]
       (.write document output-stream)
       (io/input-stream (.toByteArray output-stream))))))
