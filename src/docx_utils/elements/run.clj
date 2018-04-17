(ns docx-utils.elements.run
  (:require [clojure.tools.logging :as log])
  (:import (org.apache.poi.xwpf.usermodel XWPFRun XWPFParagraph TextSegement)
           (org.openxmlformats.schemas.wordprocessingml.x2006.main STHighlightColor$Enum)))

(defn- configure-run [^XWPFRun run {:keys [text pos bold highlight-color]
                                    :or   {text ""
                                           pos 0
                                           highlight-color "none"}}]
  (.setText run (if (string? text) text (str text)) pos)
  (.setBold run (or bold false))
  (-> run (.getCTR) (.addNewRPr) (.addNewHighlight) (.setVal (STHighlightColor$Enum/forString highlight-color))))

(defn- boolean? [value]
  (when (or (true? value) (false? value))
    true))

(defmulti set-run (fn [run value] (cond
                                     (string? value) :string
                                     (number? value) :number
                                     (map? value) :map
                                     (boolean? value) :boolean
                                     :else :not-supported)))

(defmethod set-run :string [run value]
  (configure-run run {:text value}))

(defmethod set-run :number [run value]
  (configure-run run {:text (str value)}))

(defmethod set-run :boolean [run value]
  (configure-run run {:text (str value)}))

(defmethod set-run :map [run value]
  (configure-run run value))

(defmethod set-run :default [run value]
  (log/warnf "Not supported value: %s" value))

(defn run-ids-to-text
  [^XWPFParagraph par run-range]
  (let [runs (.getRuns par)]
    (reduce
     (fn [string run-id]
       (str string (.getText (.get runs run-id) 0)))
     ""
     run-range)))

(defn remove-rest-runs!
  [^XWPFParagraph par run-range]
  (reduce
   (fn [_ run-id]
     (.removeRun par run-id))
   nil
   (reverse (rest run-range))))

(defn find-first-found-run [^XWPFParagraph par ^TextSegement found-segment]
  (.get (.getRuns par)
        (.getBeginRun found-segment)))

(defn run-id-range [^TextSegement found-segment]
  (range
   (.getBeginRun found-segment)
   (inc (.getEndRun found-segment))))

(defn merge-runs!
  [^XWPFParagraph par ^TextSegement found-segment]
  (let [run-ids (run-id-range found-segment)]
    (.setText
     (find-first-found-run par found-segment)
     (run-ids-to-text par run-ids) 0)
    (remove-rest-runs! par (run-id-range found-segment))))
