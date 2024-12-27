(defun c:TextOutline ()
  (prompt "\nPilih teks untuk ditambahkan garis tepi.")
  (setq ss (ssget '((0 . "TEXT,MTEXT")))) ; Memilih objek teks dan mtext
  (if ss
    (progn
      (setq count (sslength ss)) ; Hitung jumlah objek yang dipilih
      (setq i 0) ; Inisialisasi indeks

      ; Meminta input faktor offset dari pengguna (misalnya 0.8 atau 1.5)
      (setq offset-factor (getreal "\nMasukkan faktor offset untuk garis tepi (misal 0.8, 1.5): "))
      (if (not offset-factor) ; Jika tidak ada input yang valid
        (setq offset-factor 0.8) ; Tetapkan nilai default 0.8 jika input tidak valid
      )

      ; Meminta input radius fillet dari pengguna
      (setq fillet-radius (getreal "\nMasukkan radius fillet untuk sudut (misal 0.5, 1.0): "))
      (if (not fillet-radius) ; Jika tidak ada input yang valid
        (setq fillet-radius 0.5) ; Tetapkan nilai default 0.5 jika input tidak valid
      )

      (while (< i count)
        (setq text-entity (ssname ss i)) ; Ambil entitas teks satu per satu
        (setq text-data (entget text-entity)) ; Mendapatkan data teks
        (setq text-position (cdr (assoc 10 text-data))) ; Posisi teks
        (setq text-height (cdr (assoc 40 text-data))) ; Tinggi teks

        ; Perhitungan lebar teks
        (if (= (cdr (assoc 0 text-data)) "TEXT")
          (setq text-width (* (strlen (cdr (assoc 1 text-data))) (/ text-height 2))) ; Perkiraan lebar teks untuk TEXT
          (setq text-width (cdr (assoc 41 text-data))) ; Lebar untuk MTEXT
        )
        
        ; Menghitung offset berdasarkan faktor dan tinggi teks
        (setq offset (* text-height offset-factor)) ; Offset proporsional terhadap tinggi teks

        ; Hitung titik sudut kotak garis tepi
        (setq p1 (list (- (car text-position) offset) (- (cadr text-position) offset))) ; Kiri bawah
        (setq p2 (list (+ (car text-position) text-width offset) (- (cadr text-position) offset))) ; Kanan bawah
        (setq p3 (list (+ (car text-position) text-width offset) (+ (cadr text-position) text-height offset))) ; Kanan atas
        (setq p4 (list (- (car text-position) offset) (+ (cadr text-position) text-height offset))) ; Kiri atas

        ; Membuat polyline tertutup
        (command "_PLINE" p1 p2 p3 p4 "C") ; "C" untuk menutup polyline

        ; Terapkan fillet pada sudut-sudut polyline
        (command "_FILLET" "R" fillet-radius) ; Set radius fillet
        (command "_FILLET" (entlast) (entlast))
        
        ; Membuat ulang sudut yang melengkung dengan menghapus segmen tajam
        (mapcar
          (lambda (pt)
            (command "_FILLET" "P" pt))
          (list p1 p2 p3 p4))

        (setq i (1+ i)) ; Increment indeks
      )
      (prompt "\nGaris tepi berhasil dibuat dengan sudut melengkung untuk semua teks yang dipilih.")
    )
    (prompt "\nTidak ada teks yang dipilih.")
  )
  (princ)
)
