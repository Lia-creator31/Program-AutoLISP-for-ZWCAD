(defun c:FAW ()
  (vl-load-com) ; Memastikan bahwa Visual LISP COM tersedia
  (setq ss (ssget "X" '((8 . "100")))) ;
  (if ss
    (progn
      (setq count (sslength ss))
      (if (> count 0)
        (progn
          (repeat count
            (setq ent (ssname ss (setq count (1- count)))) 
            (command "_.explode" ent) 
          )
          (princ (strcat "\n" (itoa (sslength ss)) " object(s) on layer successfully executed."))
        )
        (princ "\nNo objects found.")
      )
    )
    (princ "\nNo objects found on layer.")
  )
  (princ) 
)
(princ)