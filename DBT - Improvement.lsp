(defun LM:GetAttributeValue ( block tag )
    (vl-some
        (function
            (lambda ( attrib )
                (if (eq tag (vla-get-Tagstring attrib))
                    (vla-get-TextString attrib)
                )
            )
        )
        (vlax-invoke block 'GetAttributes)
    )
)

(defun c:dbt (/ a b c d ss ent blkdata blk_pt b1 c1 d1 choice distance-name distance-code distance-swbs offset)
    (setq a "CIRCUIT")
    (setq b "ROOM_NAME")
    (setq c "ROOM_CODE")
    (setq d "SWBS")
    
    ;; Offset untuk menggeser teks ke kiri
    (setq offset -600) ;; Ubah nilai ini untuk mengatur seberapa jauh teks digeser ke kiri

    ;; Meminta jarak dari pengguna
    (setq distance-name (getreal "\nEnter the distance for ROOM_NAME: "))
    (setq distance-code (getreal "\nEnter the distance for ROOM_CODE: "))
    (setq distance-swbs (getreal "\nEnter the distance for SWBS: "))

    ;; Meminta pilihan dari pengguna
    (setq choice (getstring "\nDo you want to display attributes by [Select/Layer]? "))

    (cond
        ((equal (strcase choice) "SELECT")
            ;; Mengambil objek yang dipilih
            (setq ss (ssget '((0 . "INSERT")))) ; Mengambil semua objek blok
            (if ss
                (progn
                    (setq n (sslength ss))
                    (repeat n
                        (setq ent (ssname ss (setq n (1- n)))) ; Iterasi setiap objek dalam selection set
                        (setq blkdata (entget ent))
                        (setq blk_pt (cdr (assoc 10 blkdata)))

                        ;; Menampilkan ROOM_NAME di atas ROOM_CODE
                        (setq b1 (LM:GetAttributeValue (vlax-ename->vla-object ent) (strcase b)))
                        (command "_text" (list (+ (car blk_pt) offset) (+ distance-name (cadr blk_pt))) 200 "" b1 "")

                        ;; Menampilkan ROOM_CODE di atas objek
                        (setq c1 (LM:GetAttributeValue (vlax-ename->vla-object ent) (strcase c)))
                        (command "_text" (list (+ (car blk_pt) offset) (+ distance-code (cadr blk_pt))) 200 "" c1 "")


                        ;; Menampilkan SWBS di bawah objek
                        (setq d1 (LM:GetAttributeValue (vlax-ename->vla-object ent) (strcase d)))
                        (command "_text" (list (+ (car blk_pt) offset) (+ (- distance-swbs) (cadr blk_pt))) 200 "" d1 "")
                    )
                )
                (princ "\nNo objects selected.")
            )
        )
        ((equal (strcase choice) "LAYER")
            ;; Mengambil semua objek pada layer "FDS"
            (setq ss (ssget "X" '((8 . "FDS"))))
            (if ss
                (progn
                    (setq n (sslength ss))
                    (repeat n
                        (setq ent (ssname ss (setq n (1- n)))) ; Iterasi setiap objek dalam selection set
                        (setq blkdata (entget ent))
                        (setq blk_pt (cdr (assoc 10 blkdata)))

                        ;; Menampilkan ROOM_NAME di atas ROOM_CODE
                        (setq b1 (LM:GetAttributeValue (vlax-ename->vla-object ent) (strcase b)))
                        (command "_text" (list (+ (car blk_pt) offset) (+ distance-name (cadr blk_pt))) 200 "" b1 "")

                        ;; Menampilkan ROOM_CODE di atas objek
                        (setq c1 (LM:GetAttributeValue (vlax-ename->vla-object ent) (strcase c)))
                        (command "_text" (list (+ (car blk_pt) offset) (+ distance-code (cadr blk_pt))) 200 "" c1 "")

                        ;; Menampilkan SWBS di bawah objek
                        (setq d1 (LM:GetAttributeValue (vlax-ename->vla-object ent) (strcase d)))
                        (command "_text" (list (+ (car blk_pt) offset) (+ (- distance-swbs) (cadr blk_pt))) 200 "" d1 "")
                    )
                )
                (princ "\nNo objects found on layer 'FDS'.")
            )
        )
        (t (princ "\nInvalid choice. Please enter 'Select' or 'Layer'."))
    )
    (princ)
)