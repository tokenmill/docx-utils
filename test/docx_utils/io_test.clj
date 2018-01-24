(ns docx-utils.io-test
  (:require [clojure.test :refer :all]
            [clojure.java.io :as io]
            [docx-utils.io :as docx-io])
  (:import (java.io File)
           (org.apache.poi.xwpf.usermodel XWPFDocument)))

(deftest docx-transformations-test
  (testing "Testing if template can be loaded from memory."

    (with-open [input-data (java.io.FileInputStream. (io/file (io/resource "template-1.docx")))]
      (is (= XWPFDocument (type (docx-io/load-template-from-memory input-data)))))

    ;; Exception on bad input bytes
    (is (thrown? Exception
                 (docx-io/load-template-from-memory
                  (java.io.ByteArrayInputStream. (byte-array [])))))))
