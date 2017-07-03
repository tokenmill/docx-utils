(ns docx-utils.io
  (:require [clojure.tools.logging :as log]
            [clojure.java.io :as io])
  (:import (org.apache.poi.xwpf.usermodel XWPFDocument)
           (org.apache.poi POIXMLDocument)
           (java.io File)))

(defn temp-output-file []
  (.getAbsolutePath
    (doto (File/createTempFile "output-" ".docx")
      (.deleteOnExit))))

(defn empty-template []
  (XWPFDocument.))

(defn ^XWPFDocument load-template
  "Loads the template file. If path is nil then creates an empty document.
  If template file is not nil but the file does not exist then throws an exception."
  [^String template-file-path]
  (when (and (not (nil? template-file-path))
             (not (.exists (io/as-file template-file-path))))
    (throw (Exception. (str "Template file " template-file-path " does not exits."))))
  (if (nil? template-file-path)
    (empty-template)
    (XWPFDocument. (POIXMLDocument/openPackage template-file-path))))

(defn store
  [^XWPFDocument document output-file-path]
  (when-not (.exists (io/as-file output-file-path))
    (throw (Exception. (str "Output file " output-file-path " does not exits."))))
  (log/debugf "Writing the transformed template document to the output file '%s'." output-file-path)
  (with-open [o (io/output-stream output-file-path)]
    (.write document o)))
