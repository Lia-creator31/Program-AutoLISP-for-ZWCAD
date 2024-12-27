(defun c:MLY ( / ent ss count color lw layerName)
    (setq count 0)
    (setq layerName "GARIS MATERIAL")
    (prompt "\nProses deteksi polyline warna hijau")

    (if (not (tblsearch "layer" layerName))
        (progn
            (command "_.-layer" "_M" layerName "_C" "2" "" "_L" "Continuous" "" "")
            (prompt (strcat "\nLayer baru dibuat: " layerName))
        )
    )

    (setq ss (ssget "_X" '((0 . "LWPOLYLINE") (62 . 3))))
    (if ss
        (progn
            (setq count (sslength ss))
            (prompt (strcat "\nJumlah polyline hijau terdeteksi: " (itoa count)))
            (repeat count
                (setq ent (ssname ss (setq count (1- count))))
                (setq color (entget ent))
                (entmod (subst (cons 62 2) (assoc 62 color) color))
                (entupd ent)
            )
            (prompt "\nPolyline hijau berhasil diubah menjadi kuning")
        )
    (prompt "\nTidak ada polyline hijau yang terdeteksi")
    )

    (setq ss (ssget "_X" '((0 . "LWPOLYLINE") (62 . 2))))
    (if ss
        (progn
            (setq count (sslength ss))
            (prompt (strcat "\nJumlah polyline kuning terdeteksi: " (itoa count)))
            (repeat count 
                (setq ent (ssname ss (setq count (1- count))))
                (setq color (entget ent))
                (setq lw (assoc 370 color))
                (if lw
                    (entmod (subst (cons 370 40) lw color))
                    (entmod (append color (list (cons 370 40))))
                )
                (entupd ent)
                (entmod (subst (cons 8 layerName) (assoc 8 color) color))
                (entupd ent)
            )
            (prompt "\nLineweight polyline kuning berhasil di pertebal")
        )
        (prompt "\nTidak ada polyline kuning yang terdeteksi")
    )
    (princ)
)