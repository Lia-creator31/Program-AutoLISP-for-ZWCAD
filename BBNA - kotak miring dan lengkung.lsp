(defun remove-duplicates (lst)
    "delete elemen"
    (if lst
        (cons (car lst) (remove-duplicates (vl-remove (car lst) (cdr lst))))
        nil
    )
)

(defun c:BBAN ()
    (setq obj (ssget))
    (if obj
        (progn
            (command "_.explode" obj)

            (setq explodedEntities (ssget "P"))
            (if explodedEntities
                (progn
                    (setq poinList nil)
                    (repeat (sslength explodedEntities)
                        (setq entity (ssname explodedEntities 0))
                        (setq entData (entget entity))
                        (if (= (cdr (assoc 0 entData)) "LINE")
                            (progn
                            (setq startPt (cdr (assoc 10 entData)))
                            (setq endPt (cdr (assoc 11 entData)))
                            (setq poinList (append poinList (list startPt endPt)))
                            )                      
                        )
                        (ssdel entity explodedEntities)
                    )

                    (setq poinList (remove-duplicates poinList))
                    (if (and poinList (> (length poinList) 2))
                        (progn
                            (command "_.pline")
                            (foreach pt poinList
                                (command pt)
                            )
                            (command "CLOSE")

                        )
                    )
                )
        )
    )


)
    (princ)
)