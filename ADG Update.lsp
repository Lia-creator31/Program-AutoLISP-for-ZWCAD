(defun LM:GetAttributeValue (block tag)
  ;; © Lee Mac 2010
  ;BY FATHURROJI 21-12-2023
  (vl-some
    (function
      (lambda (attrib)
        (if (eq (strcase tag) (strcase (vla-get-Tagstring attrib)))
            (vla-get-TextString attrib)
          )
        )
      )
    (vlax-invoke block 'GetAttributes)
  )
)

(defun c:adg (/ a b c d ent en1 blk blk_pt b1 c1 d1 p1 p2 p3 p4 p5 ss i mode dist-sid dist-nam dist-rc base-x offset-x)
  (setq b (strcase "sid"))
  (setq c (strcase "nam"))
  (setq d (strcase ".rc"))

  ;=====================
  ; Input jarak vertikal
  (setq dist-sid (getreal "Masukkan jarak vertikal untuk SID: "))
  (setq dist-nam (getreal "Masukkan jarak vertikal untuk NAM: "))
  (setq dist-rc (getreal "Masukkan jarak vertikal untuk .RC: "))

  ; Offset tambahan untuk X (sumbu horizontal)
  (setq offset-x -400) ; Sesuaikan nilai ini sesuai kebutuhan

  ; Pilih mode: layer atau select
  (initget 1 "LAYER SELECT")
  (setq mode (getkword "Pilih mode [LAYER/SELECT]: "))

  (cond
    ((eq mode "LAYER")
     ; Cari semua objek di layer "PENETRATION"
     (setq ss (ssget "X" '((8 . "PENETRATION"))))
     (if ss
       (progn
         ; Iterasi semua entitas yang ditemukan di layer "PENETRATION"
         (setq i 0)
         (repeat (sslength ss)
           (setq en1 (ssname ss i))
           (setq blk (vlax-ename->vla-object en1))

           ; Cek apakah objek adalah block reference
           (if (and blk (vlax-property-available-p blk 'HasAttributes))
             (progn
               ; Ambil atribut NAM untuk menentukan base-x
               (setq c1 (LM:GetAttributeValue blk (strcase c)))
               (setq blk_pt (vlax-get blk 'InsertionPoint)) ; Mendapatkan posisi block
               (setq base-x (+ (car blk_pt) offset-x)) ; Tambahkan offset ke base-x

               ; Ambil atribut SID
               (setq b1 (LM:GetAttributeValue blk (strcase b)))
               (setq p3 (list base-x (+ (cadr blk_pt) dist-sid))) ; Gunakan base-x untuk SID
               (command "_text" p3 150 "" b1 "")

               ; Tampilkan NAM (tetap di base-x)
               (setq p4 (list base-x (+ (cadr blk_pt) dist-nam)))
               (command "_text" p4 150 "" c1 "")

               ; Ambil atribut .RC
               (setq d1 (LM:GetAttributeValue blk (strcase d)))
               (setq p5 (list base-x (+ (cadr blk_pt) dist-rc))) ; Gunakan base-x untuk .RC
               (command "_text" p5 150 "" d1 "")
             )
           )
           (setq i (1+ i))
         )
       )
       (alert "No blocks found on layer PENETRATION."))
    )

    ((eq mode "SELECT")
     ; Multiple Select Objects
     (setq ss (ssget)) ; Biarkan pengguna memilih objek di layar
     (if ss
       (progn
         ; Iterasi semua entitas yang dipilih
         (setq i 0)
         (repeat (sslength ss)
           (setq en1 (ssname ss i))
           (setq blk (vlax-ename->vla-object en1))

           ; Cek apakah objek adalah block reference
           (if (and blk (vlax-property-available-p blk 'HasAttributes))
             (progn
               ; Ambil atribut NAM untuk menentukan base-x
               (setq c1 (LM:GetAttributeValue blk (strcase c)))
               (setq blk_pt (vlax-get blk 'InsertionPoint)) ; Mendapatkan posisi block
               (setq base-x (+ (car blk_pt) offset-x)) ; Tambahkan offset ke base-x

               ; Ambil atribut SID
               (setq b1 (LM:GetAttributeValue blk (strcase b)))
               (setq p3 (list base-x (+ (cadr blk_pt) dist-sid))) ; Gunakan base-x untuk SID
               (command "_text" p3 150 "" b1 "")

               ; Tampilkan NAM (tetap di base-x)
               (setq p4 (list base-x (+ (cadr blk_pt) dist-nam)))
               (command "_text" p4 150 "" c1 "")

               ; Ambil atribut .RC
               (setq d1 (LM:GetAttributeValue blk (strcase d)))
               (setq p5 (list base-x (+ (cadr blk_pt) dist-rc))) ; Gunakan base-x untuk .RC
               (command "_text" p5 150 "" d1 "")
             )
           )
           (setq i (1+ i))
         )
       )
       (alert "No objects selected."))
    )

    (t (alert "Invalid mode selected."))
  )

  (princ)
)
