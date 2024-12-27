(defun c:RT0 ()
    (vl-load-com)
    (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
    (setq modelSpace (vla-get-ModelSpace doc))
    (setq fixedTexts '())
    (setq tolerance 0.5)

    (defun isRotation180 (rotation)
        (<= (abs (- rotation pi)) tolerance)
    )

    (vlax-for entity modelSpace
        (if (and (eq (vla-get-ObjectName entity) "AcDbText"))
            (progn
                (setq rotation (vla-get-Rotation entity))
                (if (isRotation180 rotation)
                    (progn
                        (vla-put-Rotation entity 0)
                        (setq fixedTexts (cons entity fixedTexts))
                    )
                )
            )
        )
    )

    (if fixedTexts
        (progn
            (princ "\nTeks dengan rotasi 180  derajat telah diubah menjadi 0 derajat")
            (foreach text fixedTexts
                (princ (strcat "\nIsi Teks: " (vla-get-TextString text)))
            )
        )
        (princ "\Tidak ada teks dengan rotasi 180 derajat")
    )
    (princ)
)