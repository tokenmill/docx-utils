(ns docx-utils.elements.run
  (:require [clojure.tools.logging :as log])
  (:import (org.apache.poi.xwpf.usermodel XWPFRun)
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
