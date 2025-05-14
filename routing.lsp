(defun C:GetLayerPoly (/ ss blkref blkpt poly layername i msg seen_layers unique_layer_count blkname acDoc ms acTable tablept row secondColumnText systemText)
  (vl-load-com)
  ; Inisialisasi ActiveX
  (setq acDoc (vla-get-ActiveDocument (vlax-get-acad-object)))
  (setq ms (vla-get-ModelSpace acDoc))
  
  ; Loop untuk memilih Block Reference sampai pengguna menekan Esc
  (while (progn
           (princ "\nPilih Block Reference (tekan Esc untuk keluar): ")
           (setq ss (ssget ":S" '((0 . "INSERT")))))
    (setq blkref (ssname ss 0))
    (setq blkpt (cdr (assoc 10 (entget blkref))))
    (setq blkname (cdr (assoc 2 (entget blkref))))
    ; Pilih POLYLINE atau LWPOLYLINE yang melewati titik blkpt
    (setq poly (ssget "C" blkpt blkpt '((-4 . "<OR") (0 . "POLYLINE") (0 . "LWPOLYLINE") (-4 . "OR>"))))
    (if poly
      (progn
        (setq msg (strcat "Daftar circuit yang melewati Blocks " blkname " : \n"))
        (setq seen_layers '())
        (setq unique_layer_count 0)
        (setq i 0)
        (repeat (sslength poly)
          (setq layername (cdr (assoc 8 (entget (ssname poly i)))))
          (if (not (member layername seen_layers))
            (progn
              (setq seen_layers (cons layername seen_layers))
              (setq unique_layer_count (1+ unique_layer_count))
              (setq msg (strcat msg "\nCircuit " (itoa unique_layer_count) ": " layername))
            )
          )
          (setq i (1+ i))
        )
        (if (> unique_layer_count 0)
          (progn
            (alert msg)
            (setq tablept (getpoint "\nKlik lokasi untuk meletakkan tabel: "))
            ; Minta input teks untuk kolom System
            (initget 1) ; Pastikan input tidak kosong
            (setq systemText (getstring T "\nMasukkan teks untuk kolom System: "))
            (if tablept
              (progn
                (vla-AddText ms
                            (strcat "Block: " blkname)
                            (vlax-3d-point (list (+ (car tablept) 150.0) (+ (cadr tablept) 10.0) 0.0))
                            5.0)
                ; Buat tabel dengan jumlah baris sesuai jumlah layer unik
                (setq acTable (vla-AddTable ms
                                            (vlax-3d-point tablept)
                                            (+ 1 unique_layer_count) ; Jumlah baris (header + layer unik)
                                            6 ; Jumlah kolom
                                            10.0 ; Tinggi baris
                                            50.0 ; Lebar kolom sementara
                              ))
                (if acTable
                  (progn
                    (princ "\nDebug: Tabel berhasil dibuat")
                    ; Remark: Cells should not be merged by default, but unmerge to be safe
                    (vla-UnmergeCells acTable 0 0 0 5) ; 5 karena ada 6 kolom (0-based index)
                    ; Atur lebar kolom
                    (vla-SetColumnWidth acTable 0 50.0) ; Kolom No.
                    (vla-SetColumnWidth acTable 1 50.0) ; Kolom ODH
                    (vla-SetColumnWidth acTable 2 70.0) ; Kolom Circuit
                    (vla-SetColumnWidth acTable 3 70.0) ; Kolom System
                    (vla-SetColumnWidth acTable 4 70.0) ; Kolom Cable
                    (vla-SetColumnWidth acTable 5 70.0) ; Kolom Number
                    ; Set teks header
                    (vla-SetText acTable 0 0 "No.")
                    (vla-SetText acTable 0 1 "ODH")
                    (vla-SetText acTable 0 2 "Circuit")
                    (vla-SetText acTable 0 3 "System")
                    (vla-SetText acTable 0 4 "Cable")
                    (vla-SetText acTable 0 5 "Cat")
                    
                    ; Debug header
                    (princ (strcat "\nDebug: Header kolom 0: " (vla-GetText acTable 0 0)))
                    (princ (strcat "\nDebug: Header kolom 1: " (vla-GetText acTable 0 1)))
                    (princ (strcat "\nDebug: Header kolom 2: " (vla-GetText acTable 0 2)))
                    (princ (strcat "\nDebug: Header kolom 3: " (vla-GetText acTable 0 3)))
                    (princ (strcat "\nDebug: Header kolom 4: " (vla-GetText acTable 0 4)))
                    (princ (strcat "\nDebug: Header kolom 5: " (vla-GetText acTable 0 5)))
                    ; Isi data tabel hanya untuk layer unik
                    (setq row 1)
                    (foreach layer seen_layers
                      (vla-SetText acTable row 0 (itoa row)) ; Kolom No.

                      ;; Kolom ODH
                      (setq lastUnderscorePos (vl-string-position (ascii "_") layer nil T))
                      (if lastUnderscorePos
                        (progn
                          (setq odhText (substr layer (+ lastUnderscorePos 2)))
                          (setq secondLastUnderscorePos (vl-string-position (ascii "_") odhText nil T))
                          (if secondLastUnderscorePos
                            (setq odhText (substr odhText (+ secondLastUnderscorePos 2)))
                          )
                          (vla-SetText acTable row 1 odhText)
                        )
                        (vla-SetText acTable row 1 "N/A")
                      )

                      ;; Kolom Circuit
                      (setq firstUnderscorePos (vl-string-position (ascii "_") layer))
                      (if firstUnderscorePos
                        (setq circuitText (substr layer 1 firstUnderscorePos))
                        (setq circuitText layer)
                      )
                      (vla-SetText acTable row 2 circuitText)

                      ;; Kolom System
                      (vla-SetText acTable row 3 systemText) ; Isi dengan input pengguna

                      ;; Kolom Cable
                      (if firstUnderscorePos
                        (progn
                          (setq afterFirstUnderscore (substr layer (+ firstUnderscorePos 2)))
                          (setq secondUnderscorePos (vl-string-position (ascii "_") afterFirstUnderscore))
                          (if secondUnderscorePos
                            (progn
                              (setq afterSecondUnderscore (substr afterFirstUnderscore (+ secondUnderscorePos 2)))
                              (setq thirdUnderscorePos (vl-string-position (ascii "_") afterSecondUnderscore))
                              (if thirdUnderscorePos
                                (setq segmentText (substr afterSecondUnderscore 1 thirdUnderscorePos))
                                (setq segmentText "N/A")
                              )
                            )
                            (setq segmentText "N/A")
                          )
                        )
                        (setq segmentText "N/A")
                      )
                      (vla-SetText acTable row 4 segmentText)

                      ;; Kolom Number
                      (if firstUnderscorePos
                        (progn
                          (setq afterFirstUnderscore (substr layer (+ firstUnderscorePos 2)))
                          (setq secondUnderscorePos (vl-string-position (ascii "_") afterFirstUnderscore))
                          (if secondUnderscorePos
                            (progn
                              (setq secondColumnText (substr afterFirstUnderscore 1 secondUnderscorePos))
                              (setq secondColumnText (vl-string-subst "" "(" secondColumnText))
                              (setq secondColumnText (vl-string-subst "" ")" secondColumnText))
                            )
                            (setq secondColumnText "N/A")
                          )
                        )
                        (setq secondColumnText "N/A")
                      )
                      (vla-SetText acTable row 5 secondColumnText)

                      (setq row (1+ row))
                    )
                  )
                  (alert "\nGagal membuat tabel."))
              )
              (alert "\nLokasi tabel tidak dipilih."))
          )
          (alert "\nTidak ada layer unik yang ditemukan dari Polyline atau Lightweight Polyline."))
      )
      (alert "\nTidak ada Polyline atau Lightweight Polyline yang melewati Block Reference."))
  )
  (princ "\nSelesai.")
  (princ)
)

