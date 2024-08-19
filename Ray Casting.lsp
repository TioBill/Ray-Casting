(defun c:ExportartPontos ( / *error* closedPolylines tipoLista pointsLista 
                       pointsInside myXL
                      AH:colocarNaCelula AH:abrirExcel checkColor OSM:DPL
                      LM:str->lst replace_word
                      currentPolyline coordenadasPolylineSeparadas coordenadasCurrentPolyline a
                      currentSelection currentObject currentLayer
                      )
  (defun *error* (msg)
    (or
      (wcmatch (strcase msg) "*BREAK, *CANCEL*, *EXIT*")
      (alert (strcat "ERROR: " msg "**"))
    )
  )
  
  (defun LM:str->lst ( str del / pos )
      (if (setq pos (vl-string-search del str))
          (cons (substr str 1 pos) (LM:str->lst (substr str (+ pos 1 (strlen del))) del))
          (list str)
      )
  )
  
  (defun replace_word (word remove new)
    (while (vl-string-search remove word)
      (setq word (vl-string-subst new remove word)) 
    )
    word
  )
  
  (defun OSM:DPL (/ fn fname linha linha1 n pt)
    (setvar 'cmdecho 0)

    (entmake '((0 . "LAYER") (100 . "AcDbSymbolTableRecord") (100 . "AcDbLayerTableRecord") (2 . "LIGACAO") (62 . 3) (70 . 0)))
    
    
    (if (setq fn (getfiled "Arquivo a abrir" (getvar 'dwgprefix) "CSV;TXT" 4))
      (progn
        (setq fn (strcase fn))
        (if (= (substr fn (- (strlen fn) 3))".CSV")(setq separator ";")(setq separator "\t"))
        (setq fname (open fn "r"))
        (setq linha (read-line fname))
      
        
        (while linha
          (setq linha1 (LM:str->lst (replace_word linha "," ".") separator))
          (setq n (atoi (car linha1))) 
          (setq pt (list (atof (cadr linha1))(atof (caddr linha1))))
          
          (if (< (distance pt (list 0 0 0)) 50)
            (setq n 0)
          )
          
          (repeat n (entmakex (list (cons 0 "POINT")(cons 8 "LIGACAO")(cons 10 pt))))
          (setq linha (read-line fname))
        )
        (close fname)
      )
    )
    (princ)
  )
  
  (defun AH:colocarNaCelula (cellname val1 / myRange)
    (setq myRange (vlax-get-property (vlax-get-property myXL "ActiveSheet") "Range" cellname))
    (vlax-put-property myRange 'Value2 val1)
  )
  
  (defun AH:abrirExcel()
    (if (null (setq myXL (vlax-get-object "Excel.Application")))
      (setq myXL (vlax-get-or-create-object "excel.Application"))
    )

    (vla-put-visible myXL :vlax-false)
    (vlax-put-property myXL 'ScreenUpdating :vlax-true)
    (vlax-put-property myXL 'DisplayAlerts :vlax-true)
    
    (vlax-invoke-method (vlax-get-property myXl 'WorkBooks) 'Add) 
  )
  
  
  (initget 1 "Sim Nao")
  
  (if (eq (getkword "\nDeseja importar pontos? [Sim/Nao]") "Sim")
    (OSM:DPL)
  )
  
  (setq closedPolylines (ssget "_A" '((0 . "LWPOLYLINE"))))
  
  (AH:abrirExcel)
  
  (AH:colocarNaCelula "A1" "Layer")
  (AH:colocarNaCelula "B1" "Quantidade")
  
  (setq tipoLista (list))
  (setq pointsLista (list))

  (repeat (setq contador (sslength closedPolylines))
    (setq currentPolyline (entget (ssname closedPolylines (setq contador (1- contador)))))
    
    (setq tipoLista (append tipoLista (list (vla-get-layer (vlax-ename->vla-object (ssname closedPolylines contador))))))
    
    (setq coordenadasCurrentPolyline (member (assoc 10 currentPolyline) currentPolyline))

    (setq coordenadasPolylineSeparadas nil)
    
    ;; Separate coordinates per point
    (mapcar '(lambda (a)
              (if (= (car a) 10)
                (setq coordenadasPolylineSeparadas (cons (cdr a) coordenadasPolylineSeparadas))
              )         
             )
      coordenadasCurrentPolyline
    )
    
    (setq pointsInside 0)
  
    (if (> (length coordenadasCurrentPolyline) 1)
      (progn
        (setq currentSelection (ssget "_CP" coordenadasPolylineSeparadas))
        
        (repeat (setq n (sslength currentSelection))
          (setq currentObject (ssname currentSelection (setq n (1- n))))
          (setq currentLayer (vla-get-layer (vlax-ename->vla-object currentObject)))
          
          (if (= currentLayer "LIGACAO")
            (setq pointsInside (1+ pointsInside))
          )
        ) 
      )
    )
    
    (setq pointsLista (append pointsLista (list pointsInside)))
    
    (AH:colocarNaCelula (strcat "A" (rtos (+ (length tipoLista) 1) 2 0)) (last tipoLista))
    (AH:colocarNaCelula (strcat "B" (rtos (+ (length pointsLista) 1) 2 0)) (last pointsLista))
  )

  (vla-put-visible myXL :vlax-true)
  (vlax-release-object myXL)
  
  (princ)
)

(vl-load-com)

(alert "Lisp Carregada! Digite \"ExportartPontos\" para comecar!\n\nCaso ja existam pontos no .DWG, verifique que os pontos estao na layer \"LIGACAO\".")

