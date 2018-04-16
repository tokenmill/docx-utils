(ns docx-utils.replace-test
  (:require [clojure.test :refer :all]
            [clojure.java.shell :as shell]
            [clojure.java.io :as io]
            [docx-utils.io :as docx-io]
            [docx-utils.core :as docx]
            [docx-utils.schema :as schema]))

(deftest replace-test
  (testing "Testing if transformation replaces placeholder in the template document with plain text."
    (let [output-file-path (docx/transform (.getPath (io/resource "template-2-replace-placeholder.docx"))
                                           [(schema/transformation :replace-text "PLACEHOLDER" "Replaced with plaint text.")])]
      (shell/sh "timeout" "5s" "libreoffice" "--norestore" output-file-path)))

  (testing "Testing if transformation replaces placeholder in the template document with a number."
    (let [output-file-path (docx/transform (.getPath (io/resource "template-2-replace-placeholder.docx"))
                                           [(schema/transformation :replace-text "PLACEHOLDER" 12345)])]
      (shell/sh "timeout" "5s" "libreoffice" "--norestore" output-file-path)))

  (testing "Testing if transformation replaces placeholder in the template document with inline plain text."
    (let [output-file-path (docx/transform (.getPath (io/resource "template-2-replace-placeholder.docx"))
                                           [(schema/transformation :replace-text-inline "PLACEHOLDER" "Replaced with plain inline text.")])]
      (shell/sh "timeout" "5s" "libreoffice" "--norestore" output-file-path)))

  (testing "Testing if transformation replaces placeholder in the template document with a data table."
    (let [output-file-path (docx/transform (.getPath (io/resource "template-2-replace-placeholder.docx"))
                                           [(schema/transformation :replace-table "PLACEHOLDER" [["cell 11" "cell 12" "cell 13"] ["cell 21" "cell 22" "cell 23"]])])]
      (shell/sh "timeout" "5s" "libreoffice" "--norestore" output-file-path)))

  (testing "Testing if transformation replaces placeholder in the template document with a data table where cell value is a map."
    (let [output-file-path (docx/transform (.getPath (io/resource "template-2-replace-placeholder.docx"))
                                           [(schema/transformation :replace-table "PLACEHOLDER"
                                                                   [[{:text "cell 11" :bold true} {:text "cell 12" :bold false :highlight-color "red"}]])])]
      (shell/sh "timeout" "5s" "libreoffice" "--norestore" output-file-path)))

  (testing "Testing if transformation replaces placeholder in the template document with an image."
    (let [output-file-path (docx/transform (.getPath (io/resource "template-2-replace-placeholder.docx"))
                                           [(schema/transformation :replace-image "PLACEHOLDER" (.getPath (io/resource "test-image.jpg")))])]
      (shell/sh "timeout" "5s" "libreoffice" "--norestore" output-file-path)))

  (testing "Testing if transformation replaces placeholder in the template document with a bullet list."
    (let [output-file-path (docx/transform (.getPath (io/resource "template-2-replace-placeholder.docx"))
                                           [(schema/transformation :replace-bullet-list "PLACEHOLDER" ["item 1" "item 2" "item 3"])])]
      (shell/sh "timeout" "5s" "libreoffice" "--norestore" output-file-path)))

  (testing "Testing if transformation replaces placeholder in the template document with a bullet list."
    (let [output-file-path (docx/transform (.getPath (io/resource "template-2-replace-placeholder.docx"))
                                           [(schema/transformation :replace-bullet-list "PLACEHOLDER" ["item 1" "item 2" "item 3"])])]
      (shell/sh "timeout" "5s" "libreoffice" "--norestore" output-file-path)))

  (testing "Testing if transformation replaces placeholder in the template document with a numbered list."
    (let [output-file-path (docx/transform (.getPath (io/resource "template-2-replace-placeholder.docx"))
                                           [(schema/transformation :replace-numbered-list "PLACEHOLDER" ["numbered item 1" "numbered item 2" "numbered item 3"])])]
      (shell/sh "timeout" "5s" "libreoffice" "--norestore" output-file-path)))

  (testing "Testing if transformation replaces placeholder in the template document with a bullet list where item is a map."
    (let [output-file-path (docx/transform (.getPath (io/resource "template-2-replace-placeholder.docx"))
                                           [(schema/transformation :replace-bullet-list "PLACEHOLDER"
                                                                   [{:text "item1" :bold true} {:text "item 2" :bold false :highlight-color "red"}])])]
      (shell/sh "timeout" "5s" "libreoffice" "--norestore" output-file-path)))

  (testing "Testing if tokens can be in lowercase."
    (let [output-file-path (docx/transform (.getPath (io/resource "template-4-lowercase-token.docx"))
                                           [(schema/transformation :replace-text "lowercase_token" "Lowercase tokens get replaced.")])]
      (shell/sh "timeout" "5s" "libreoffice" "--norestore" output-file-path))))

(deftest byte-stream-replace-test
  (testing "Testing if transformation replaces placeholder in the template document with plain text."
    (let [input-stream (docx/transform-input-stream
                        [(schema/transformation :replace-text "PLACEHOLDER" "Byte stream replace works.")]
                        (io/input-stream (io/resource "template-2-replace-placeholder.docx")))
          output-file-path (docx-io/temp-output-file)]
      (with-open [o (io/output-stream (io/file output-file-path))]
        (io/copy input-stream o))
      (shell/sh "timeout" "5s" "libreoffice" "--norestore" output-file-path))))

(deftest replace-multi-run-test
  (testing "multiple runs in paragraph"
    (let [output-file-path (docx/transform (.getPath (io/resource "template-5-multi-run.docx"))
                                           [(schema/transformation :replace-text-inline "name" "Replaced with plaint text.")])]
      (shell/sh "timeout" "5s" "libreoffice" "--norestore" output-file-path))))
