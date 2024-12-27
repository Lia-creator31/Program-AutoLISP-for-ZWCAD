(defun AddToGroup (key value groups)
    (if (assoc key groups)
        (setq groups (mapcar (function
            (lambda (x)
                (if (equal (car x) key)
                    (cons key (append (cdr x) (list value)))
                    x)))
            groups))
        (setq groups (cons (cons key (list value)) groups)))
    groups
)

(defun C:MBB1 (/ ss n blk blk-name csv-data groups offsetX offsetY refX refY minpt maxpt LL UR width height scaleFactor CsvPath)
    (setq A4Width 297.0
        A4Height 210.0
        offsetX 500.0
        offsetY 50.0
        refX 0.0
        refY 0.0
        csv-data '()
        groups '()
    )

    (setq CsvPath (getfiled "Pilih file CSV" "" "csv" 16))
    (if (and CsvPath (setq csv-data (ReadCSV CsvPath)))
        (progn 
            (if (setq ss (ssget "X" '((0 . "INSERT"))))
                (progn
                    (repeat (setq n (sslength ss))
                    (setq blk (ssname ss (setq n (1- n))))
                    (setq blk-name (cdr (assoc 2 (entget blk))))
                    (if (not blk-name)
                        (setq blk-name "Unknown"))
                    (foreach desc csv-data
                        (if (wcmatch blk-name (strcat desc "*"))   
                            (setq groups (AddToGroup desc blk groups)))))
                

                (foreach group groups
                ;;Tambahkan judul kelompok
                    (setq refY 0.0)
                    (foreach blk (cdr group)
                        (vla-getboundingbox (vlax-ename->vla-object blk) 'minpt 'maxpt)
                        (setq LL (vlax-safearray->list minpt)
                              UR (vlax-safearray->list maxpt))
                        (setq width (- (car UR) (car LL))
                              height (- (cadr UR) (cadr LL)))

                        (if (and (> width 0) (> height 0))
                            (progn 
                                (setq scaleX (/ A4Width width)
                                    scaleY (/ A4Height  height)
                                    scaleFactor (if (< scaleX scaleY) scaleX scaleY))

                                (command "_.scale" blk "" LL scaleFactor)

                                (command "_.move" blk "" LL (list refX refY))

                                (setq refY (+ refY height offsetY)))))
                
                (setq refX (+ refX A4Width offsetX))))
            
        (alert "Tidak ada blok yang ditemukan.")))
        (alert "CSV file tidak dapat di baca"))
        
    (princ)

)
