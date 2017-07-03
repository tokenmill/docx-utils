(ns docx-utils.core-test
  (:require [clojure.test :refer :all]
            [clojure.java.io :as io]
            [clojure.java.shell :refer [sh]]
            [docx-utils.core :as docx]
            [docx-utils.schema :refer [transformation]])
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
    (let [output-file-path (docx/transform [(transformation :append-text "appended text")])]
      (sh "timeout" "5s" "libreoffice" "--norestore" output-file-path)))

  (testing "Testing if transformation appends 1 Paragraph to the template document."
    (let [output-file-path (docx/transform [{:type :append-text :replacement "appended text as a map"}])]
      (sh "timeout" "5s" "libreoffice" "--norestore" output-file-path)))

  (testing "Testing if transformation appends 1 image to the template document."
    (let [output-file-path (docx/transform [{:type :append-image :replacement (.getPath (io/resource "test-image.jpg"))}])]
      (sh "timeout" "5s" "libreoffice" "--norestore" output-file-path)))

  (testing "Testing if transformation appends 1 table to the template document."
    (let [output-file-path (docx/transform [{:type :append-table :replacement [["cell 11" "cal 12" "cell 13"] ["cell 21" "cell 22" "cell 23"]]}])]
      (sh "timeout" "5s" "libreoffice" "--norestore" output-file-path)))

  (testing "Testing if transformation appends 1 bullet list to the template document."
    (let [output-file-path (docx/transform [{:type :append-bullet-list :replacement ["item a" "item b" "item c"]}])]
      (sh "timeout" "5s" "libreoffice" "--norestore" output-file-path)))

  (testing "Testing if transformation appends 1 numbered list to the template document."
    (let [output-file-path (docx/transform [{:type :append-numbered-list :replacement ["item 1" "item 2" "item 3"]}])]
      (sh "timeout" "5s" "libreoffice" "--norestore" output-file-path))))
