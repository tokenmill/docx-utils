(ns docx-utils.image
  (:require [clojure.java.io :as io]
            [docx-utils.constants :refer [page-width-in-emu]])
  (:import (org.apache.poi.xwpf.usermodel XWPFPicture XWPFRun Document)
           (java.awt.image BufferedImage)
           (javax.imageio ImageIO)
           (org.apache.poi.util Units)))

(defn- image-size-in-emu [width-in-pixels height-in-pixels]
  (let [width-in-emu (min page-width-in-emu (Units/pixelToEMU width-in-pixels))
        height-in-emu (min page-width-in-emu (Units/pixelToEMU height-in-pixels))]
    (if (and (not= width-in-emu page-width-in-emu) (not= height-in-emu page-width-in-emu))
      [width-in-emu height-in-emu]
      (if (= width-in-emu page-width-in-emu)
        [width-in-emu (* height-in-pixels (/ page-width-in-emu  width-in-pixels))]
        [(* width-in-pixels (/ page-width-in-emu height-in-pixels)) height-in-emu]))))

(defn ^XWPFPicture insert [^XWPFRun run image-path]
  (let [^BufferedImage bi (ImageIO/read (io/file image-path))
        [width-in-emu height-in-emu] (image-size-in-emu (.getWidth bi) (.getHeight bi))
        file-name (.getName (io/file image-path))]
    (with-open [image (io/input-stream image-path)]
      (.addPicture run
                   image
                   Document/PICTURE_TYPE_JPEG
                   file-name
                   width-in-emu
                   height-in-emu))))
