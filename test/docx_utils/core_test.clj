(ns docx-utils.core-test
  (:require [clojure.test :refer :all]
            [clojure.java.io :as io]
            [docx-utils.core :as docx])
  (:import (java.io File)))

(deftest docx-transformations-test
  (testing "Testing if transformation returns a file path of an existing file when Transformation list is nil."
    (let [template-file-path (.getPath (io/resource "template-1.docx"))
          output-file-path (docx/transform nil)]
      (is (string? output-file-path))
      (is (.exists (io/as-file output-file-path)))))

  (testing "Testing if transformation returns an exception when template file path is nil."
    (let [template-file-path (.getPath (io/resource "template-1.docx"))
          output-file-path (.getAbsolutePath (doto (File/createTempFile "output-" ".docx")
                                               (.deleteOnExit)))]
      (is (thrown? Exception (docx/transform nil nil)))
      (is (thrown? Exception (docx/transform nil [])))
      (is (string? (docx/transform template-file-path nil)))
      (is (string? (docx/transform template-file-path [])))
      (is (string? (docx/transform template-file-path output-file-path [])))
      (is (thrown? Exception (docx/transform template-file-path nil nil)))
      (is (string? (docx/transform template-file-path output-file-path nil)))
      (is (thrown? Exception (docx/transform nil nil nil))))))
