(ns docx-utils.append-test
  (:require [clojure.test :refer :all]
            [clojure.java.shell :refer [sh]]
            [clojure.java.io :as io]
            [docx-utils.core :as docx]
            [docx-utils.schema :refer [transformation]]))


(deftest append-test
  (testing "Testing if transformation appends 1 Paragraph to the template document."
    (let [output-file-path (docx/transform [(transformation :append-text "appended text")])]
      (sh "timeout" "5s" "libreoffice" "--norestore" output-file-path)))

  (testing "Testing if transformation appends 1 Paragraph to the template document."
    (let [output-file-path (docx/transform [(transformation :append-text-inline "appended inline text")])]
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