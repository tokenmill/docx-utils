(ns docx-utils.elements.run
  (:import (org.apache.poi.xwpf.usermodel XWPFRun)
           (org.openxmlformats.schemas.wordprocessingml.x2006.main STHighlightColor$Enum)))

(defn set-run [^XWPFRun run & {:keys [text pos bold highlight-color]
                               :or   {text ""
                                      pos 0
                                      highlight-color "none"}}]
  (.setText run (if (string? text) text (str text)) pos)
  (.setBold run (or bold false))
  (-> run (.getCTR) (.addNewRPr) (.addNewHighlight) (.setVal (STHighlightColor$Enum/forString highlight-color))))
