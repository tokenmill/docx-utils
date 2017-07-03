(ns docx-utils.core-test
  (:require [clojure.test :refer :all]
            [clojure.java.io :as io]
            [clojure.java.shell :refer [sh]]
            [docx-utils.core :as docx]
            [docx-utils.transformation :refer [transformation]])
  (:import (java.io File)))

(deftest docx-transformations-test
  (testing "Testing if transformation returns a file path of an existing file when Transformation list is nil."

    ;; when
    (let [output-file-path (docx/transform nil nil)]
      (is (string? output-file-path))
      (is (.exists (io/as-file output-file-path))))

    (let [template-file-path (.getPath (io/resource "template-1.docx"))
          output-file-path (docx/transform nil)]
      (is (string? output-file-path))
      (is (.exists (io/as-file output-file-path))))

    ;; When template file doesn't exits then exception
    (is (thrown? Exception (docx/transform "DOES-NOT-EXIST.docx" nil))))

  (testing "Testing if transformation appends 1 Paragraph to the template document."
    (let [template-file-path (.getPath (io/resource "template-1.docx"))
          output-file-path (docx/transform [(transformation :append-text "appended text")])]
      (is (string? output-file-path))
      (is (.exists (io/as-file output-file-path)))
      (sh "libreoffice" output-file-path)))

  (testing "Testing if transformation appends 1 Paragraph to the template document."
    (let [template-file-path (.getPath (io/resource "template-1.docx"))
          output-file-path (docx/transform [{:type :append-text :replacement "appended text as a map"}])]
      (is (string? output-file-path))
      (is (.exists (io/as-file output-file-path)))
      (sh "libreoffice" output-file-path))))
