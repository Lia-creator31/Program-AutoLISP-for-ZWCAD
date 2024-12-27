  (defun c:MTX ()
    (princ "\nPilih text yang terpisah-pisah untuk digabung")
    (setq ss (ssget))
		(if ss
			(progn 
				(setq teks "")
				(setq data-list '())
				(repeat (sslength ss)
					(setq ent (ssname ss 0))
					(setq data (entget ent))
					(if (and (assoc 1 data) (assoc 10 data))
						(progn
							(setq pos (cdr (assoc 10 data)))
							(setq text (cdr (assoc 1 data)))
							(setq data-list (cons (list pos text) data-list))
							(entdel ent)
						)
					)
					(ssdel ent ss)
				)
				(if data-list
					(progn
						(setq sorted-data (vl-sort data-list
																			(lambda (a b) (< (car (car a)) (car (car b))))))
						(foreach item sorted-data
							(setq teks (strcat teks (if (equal teks "") "" " ") (cadr item))))
						
						(setq teks (strcat "\\C1;\\H3;" teks))
						(setq start-pos (car (car (car sorted-data))))
						(command "_.MTEXT" start-pos pause teks)
						(princ "\nTeks Berhasil digabung dengan warna merah!") 
					)
					(princ "\n Tidak ada teks yang valid untuk digabung")
				)
			)
			(princ "\Tidak ada objek yang dipilih")
		)
		(princ)
	)