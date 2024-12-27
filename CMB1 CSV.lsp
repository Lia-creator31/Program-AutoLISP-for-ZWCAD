(defun C:CMB1CSV ()
    (vl-load-com)
    (defun BrowseFolder (Message / shellApp folder result)
        (setq shellApp (vlax-create-object "Shell.Application"))
        (setq folder (vlax-invoke-method shellApp 'BrowseForFolder 0 Message 0))
        (vlax-release-object shellApp)
        (if folder 
            (progn
                (Setq result (vlax-get-property folder 'Self))
                (setq result (vlax-get-property result 'Path))
                (if (/= (substr result (strlen result)) "\\")
                    (setq result (strcat result "\\"))))
        )
        result
    )

    (defun ReadCSV (filename / file line data row)
        (setq data '())
        (if (findfile filename)
            (progn
                (setq file (open filename "r"))
                (read-line file)
                (while (setq line (read-line file))
                    (setq row (vl-string->list line ";"))
                    (setq data (append data (list (cadr row))))
                )
                (close file))
            (alert (strcat "File tidak ditemukan: " filename)))
        data
    )

    (defun vl-string->list (str delim)
        (if (vl-string-search delim str)
            (cons (substr str 1 (vl-string-search delim str))
                  (vl-string->list (substr str (+ (vl-string-search delim str) 2)) delim))
            (list str)))
    
    (if (setq DirPath (BrowseFolder "Select directory to scan drawings"))
        (progn
            (setq Scrfile (strcat DirPath "tes2.scr"))
            (setq ofile (open Scrfile "w"))
            (setq DwgList (vl-directory-files DirPath "*.dxf" 1))

            (setq CsvPath (getfiled "Select CSV file for sorting" "" "csv" 16))
            (if CsvPath
                (progn
                    (setq sorted-names (ReadCSV CsvPath))
                    (if (not sorted-names)
                        (alert "CSV file is empty or invalid.")
                        (progn
                            (setq x 0)
                            (setq count 0)
                            (foreach csv-name sorted-names
                                (foreach file DwgList
                                    (if (and csv-name (wcmatch file (strcat csv-name "*")))
                                        (progn
                                            (command "-insert" (strcat DirPath file) (list x 0.0 0.0) "" "" "")
                                            (setq x (+ x 100))
                                            (setq count (+ count 1))
                                        ))))
                            (alert (strcat "Total file DXF yang berhasil dibuka: " (itoa count)))))
                )
                (alert "No CSV file selected"))))
        (princ))
    



                    