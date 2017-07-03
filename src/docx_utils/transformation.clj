(ns docx-utils.transformation)

(defrecord Transformation [type placeholder replacement])

(defn transformation
  "There are wo types of transformation: append and replace."
  ([type value]
    (->Transformation type nil value))
  ([type placeholder replacement]
    (->Transformation type placeholder replacement)))
