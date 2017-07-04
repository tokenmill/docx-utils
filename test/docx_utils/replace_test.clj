(ns docx-utils.replace-test
  (:require [clojure.test :refer :all]
            [clojure.java.shell :refer [sh]]
            [clojure.java.io :as io]
            [docx-utils.core :as docx]
            [docx-utils.schema :refer [transformation]]))

(deftest replace-test
  (testing "Testing if transformation replaces placeholder in the template document with plain text."
    (let [output-file-path (docx/transform (.getPath (io/resource "template-2-replace-placeholder.docx")) [(transformation :replace-text "${PLACEHOLDER}" "Replaced with plaint text.")])]
      (sh "timeout" "5s" "libreoffice" "--norestore" output-file-path)))

  (testing "Testing if transformation replaces placeholder in the template document with inline plain text."
    (let [output-file-path (docx/transform (.getPath (io/resource "template-2-replace-placeholder.docx")) [(transformation :replace-text-inline "${PLACEHOLDER}" "Replaced with plain inline text.")])]
      (sh "timeout" "5s" "libreoffice" "--norestore" output-file-path)))

  (testing "Testing if transformation replaces placeholder in the template document with a data table."
    (let [output-file-path (docx/transform (.getPath (io/resource "template-2-replace-placeholder.docx")) [(transformation :replace-table "${PLACEHOLDER}" [["cell 11" "cell 12" "cell 13"] ["cell 21" "cell 22" "cell 23"]])])]
      (sh "timeout" "5s" "libreoffice" "--norestore" output-file-path)))

  (testing "Testing if transformation replaces placeholder in the template document with an image."
    (let [output-file-path (docx/transform (.getPath (io/resource "template-2-replace-placeholder.docx")) [(transformation :replace-image "${PLACEHOLDER}" (.getPath (io/resource "test-image.jpg")))])]
      (sh "timeout" "5s" "libreoffice" "--norestore" output-file-path)))

  (testing "Testing if transformation replaces placeholder in the template document with a bullet list."
    (let [output-file-path (docx/transform (.getPath (io/resource "template-2-replace-placeholder.docx")) [(transformation :replace-bullet-list "${PLACEHOLDER}" ["item 1" "item 2" "item 3"])])]
      (sh "timeout" "5s" "libreoffice" "--norestore" output-file-path)))

  (testing "Testing if transformation replaces placeholder in the template document with a numbered list."
    (let [output-file-path (docx/transform (.getPath (io/resource "template-2-replace-placeholder.docx")) [(transformation :replace-numbered-list "${PLACEHOLDER}" ["numbered item 1" "numbered item 2" "numbered item 3"])])]
      (sh "timeout" "5s" "libreoffice" "--norestore" output-file-path))))
