; (defun c:outline ()
;     (vl-load-com)
;     (setq sel (ssget))
;     (if sel
;         (progn
;             (setq count (sslength sel))
;             (setq i 0)
;             (while (< i count)
;                 (setq obj (vlax-ename->vla-object (ssname sel i)))
;                 (if obj
;                     (progn
;                         (if (not (or (vl-position (vla-get-ObjectName obj) '("AcDbText" "AcDbMText")))
;                                  )
;                             (progn
;                                 (vla-getboundingbox obj 'minp 'maxp)
;                                 (setq minp (vlax-safearray->list minp))
;                                 (setq maxp (vlax-safearray->list maxp))
;                                 (setq p1 (list (car minp) (cadr minp))
;                                     p2 (list (car maxp) (cadr minp))
;                                     p3 (list (car maxp) (cadr maxp))
;                                     p4 (list (car minp) (cadr maxp)))
;                                 (command "_.PLINE" p1 p2 p3 p4 "C")
;                             )
;                             (princ (strcat "\nObjek ke-" (itoa i) "adalah teks dan dilewati."))
;                         )
;                     )
;                     (princ (strcat "\nObjek ke-" (itoa i) "tidak valid."))
;                 )
;                 (setq i (1+ i))
;             )
;             (princ "\nGaris berhasil dibuat")
;         )
;         (princ "\Tidak ada objek yang dipilih")
;     )
;     (princ)
; )

(defun c:outline ()
    (vl-load-com)
    (setq sel (ssget))
    (if sel
        (progn
            (setq count (sslength sel))
            (setq bounding-list '())
            
            ;; Loop melalui semua objek yang dipilih
            (repeat count
                (setq obj (vlax-ename->vla-object (ssname sel (setq i (1- count)))))
                (if (and obj (not (vl-position (vla-get-ObjectName obj) '("AcDbText" "AcDbMText"))))
                    (progn
                        ;; Ambil bounding box objek
                        (vla-getboundingbox obj 'minp 'maxp)
                        (setq minp (vlax-safearray->list minp))
                        (setq maxp (vlax-safearray->list maxp))
                        ;; Tambahkan ke daftar bounding box
                        (setq bounding-list
                              (append bounding-list
                                      (list (list (car minp) (cadr minp) (car maxp) (cadr maxp)))))
                    )
                )
            )

            ;; Fungsi untuk memeriksa apakah dua bounding box saling bersentuhan
            (defun bounding-box-overlap (box1 box2)
                (not (or (> (car box1) (nth 2 box2))  ; Kiri > Kanan
                         (< (nth 2 box1) (car box2))  ; Kanan < Kiri
                         (> (cadr box1) (nth 3 box2)) ; Bawah > Atas
                         (< (nth 3 box1) (cadr box2)) ; Atas < Bawah
                    )
                )
            )

            ;; Fungsi untuk membuat grup bounding box berdasarkan koneksi
            (defun group-bounding-boxes (boxes)
                (if (null boxes)
                    '()
                    (let ((groups '())) ; Daftar grup bounding box
                        (while boxes
                            (setq cur (car boxes))
                            (setq boxes (cdr boxes))
                            ;; Cari grup yang berhubungan dengan bounding box saat ini
                            (setq related-group nil)
                            (setq new-groups '())
                            (foreach group groups
                                (if (some (lambda (box) (bounding-box-overlap box cur)) group)
                                    (setq related-group (append related-group group))
                                    (setq new-groups (append new-groups (list group)))))
                            ;; Tambahkan bounding box ke grup terkait atau buat grup baru
                            (if related-group
                                (setq groups (append new-groups (list (append related-group (list cur)))))
                                (setq groups (append groups (list (list cur))))))
                        groups)))

            ;; Fungsi untuk menggabungkan bounding box dalam satu grup menjadi satu outline
            (defun merge-group (group)
                (let ((min-x (apply 'min (mapcar 'car group)))
                      (min-y (apply 'min (mapcar 'cadr group)))
                      (max-x (apply 'max (mapcar 'nth group 2)))
                      (max-y (apply 'max (mapcar 'nth group 3))))
                  (list min-x min-y max-x max-y)))

            ;; Buat grup bounding box berdasarkan koneksi
            (setq grouped-bounding-boxes (group-bounding-boxes bounding-list))

            ;; Buat outline untuk setiap grup bounding box
            (foreach group grouped-bounding-boxes
                (setq merged-box (merge-group group))
                (setq p1 (list (car merged-box) (cadr merged-box))
                      p2 (list (nth 2 merged-box) (cadr merged-box))
                      p3 (list (nth 2 merged-box) (nth 3 merged-box))
                      p4 (list (car merged-box) (nth 3 merged-box)))
                (command "_.PLINE" p1 p2 p3 p4 "C")
            )
            (princ "\nGaris tepi berhasil dibuat.")
        )
        (princ "\nTidak ada objek yang dipilih.")
    )
    (princ)
)
