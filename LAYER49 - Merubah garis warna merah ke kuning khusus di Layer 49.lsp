(defun c:Layer49 ()
    (vl-load-com)
    (setq doc(vla-get-ActiveDocument (vlax-get-acad-object)))
    (setq modelSpace (vla-get-ModelSpace doc)) 

    (vlax-for obj modelSpace
        (if (and 
                (= (vla-get-Layer obj) "49")
                (= (vla-get-Color obj) 1)
                (vlax-property-available-p obj 'Color))
            (vla-put-Color obj 60)
        )
    )
    (princ "\nProses selesai. Garis pada layer 49 dengan warna merah diubah menjadi warna kuning.")
    (princ)
)