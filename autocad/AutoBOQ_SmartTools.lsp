;;; ============================================================================
;;; AutoBOQ_SmartTools.lsp — Part-Specific AutoCAD to Excel BOQ Extractor
;;; Inspired by SmartCivilTools part-wise workflow (Footing, Column, Slab, Rooms)
;;;
;;; HOW TO USE IN AUTOCAD:
;;;   1. Type APPLOAD -> Select AutoBOQ_SmartTools.lsp -> Load
;;;   2. Run any part-specific command:
;;;        AB_FOOTING  -> Select closed footing polylines (exports Footings)
;;;        AB_COLUMN   -> Select closed column polylines  (exports Columns)
;;;        AB_SLAB     -> Select closed slab polylines    (exports Slabs)
;;;        AB_ROOMS    -> Select closed room polylines    (exports Rooms/Tiles/Paint)
;;;        AB_ALL      -> Select all closed polylines     (auto-classifies by Layer)
;;;   3. Upload the generated AutoBOQ_Takeoff.csv directly into AutoBOQ 2D
;;;      (Streamlit Tab 4) to build the complete 9-Tab Modular Excel BOQ!
;;; ============================================================================

(vl-load-com)

(defun _ab-csv-path (/ dwgdir)
  (setq dwgdir (getvar "DWGPREFIX"))
  (if (or (null dwgdir) (= dwgdir ""))
    (setq dwgdir (strcat (getenv "USERPROFILE") "\\Desktop\\"))
  )
  (strcat dwgdir "AutoBOQ_Takeoff.csv")
)

(defun _ab-ensure-header (fpath / fp)
  (if (not (findfile fpath))
    (progn
      (setq fp (open fpath "w"))
      (write-line "Part_Category,Layer,Mark,Nos,Area_m2,Perimeter_m,Length_m,Width_m,Depth_or_Height_m,Unit_Scale" fp)
      (close fp)
    )
  )
)

(defun _ab-unit-scale (/ ins)
  ;; Convert drawing units to meters (4 = mm, 5 = cm, 6 = m, 1 = inches, 2 = feet)
  (setq ins (getvar "INSUNITS"))
  (cond
    ((= ins 4) 0.001)
    ((= ins 5) 0.01)
    ((= ins 6) 1.0)
    ((= ins 1) 0.0254)
    ((= ins 2) 0.3048)
    (T 0.001) ;; default mm if unitless
  )
)

(defun _ab-classify-layer (lay / u)
  (setq u (strcase lay))
  (cond
    ((or (wcmatch u "*FOOT*") (wcmatch u "*FTG*") (wcmatch u "*FDN*")) "FOOTING")
    ((or (wcmatch u "*COL*") (wcmatch u "*PIL*")) "COLUMN")
    ((or (wcmatch u "*BEAM*") (wcmatch u "*BM*") (wcmatch u "*PLINTH*")) "BEAM")
    ((or (wcmatch u "*SLAB*") (wcmatch u "*ROOF*")) "SLAB")
    ((or (wcmatch u "*WALL*") (wcmatch u "*MAS*") (wcmatch u "*BRICK*")) "MASONRY")
    ((or (wcmatch u "*DOOR*") (wcmatch u "*WIN*") (wcmatch u "*OPEN*")) "OPENING")
    (T "ROOM")
  )
)

(defun _ab-export-selection (cat_override default_ht / ss i ent vlaobj lay area_raw perim_raw
                                                       minpt maxpt p1 p2 sc area_m2 perim_m len_m wid_m
                                                       ht fpath fp mark_prefix cnt cat)
  (setq sc (_ab-unit-scale))
  (princ (strcat "\n[AutoBOQ] Drawing unit scale to meters = " (rtos sc 2 4)))
  (setq ht (getreal (strcat "\nEnter Depth/Height in meters <" (rtos default_ht 2 2) ">: ")))
  (if (null ht) (setq ht default_ht))
  (setq mark_prefix (getstring (strcat "\nEnter Mark Prefix (e.g. F, C, S, R) <"
                                       (if cat_override (substr cat_override 1 1) "E") ">: ")))
  (if (= mark_prefix "") (setq mark_prefix (if cat_override (substr cat_override 1 1) "E")))

  (princ "\nSelect closed polylines (LWPOLYLINE / POLYLINE): ")
  (setq ss (ssget '((0 . "LWPOLYLINE,POLYLINE"))))
  (if (null ss)
    (progn (princ "\nNo polylines selected. Cancelled.") (princ))
    (progn
      (setq fpath (_ab-csv-path))
      (_ab-ensure-header fpath)
      (setq fp (open fpath "a"))
      (setq cnt 0 i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        (setq vlaobj (vlax-ename->vla-object ent))
        (setq lay (vla-get-Layer vlaobj))
        (setq cat (if cat_override cat_override (_ab-classify-layer lay)))
        (setq area_raw (vl-catch-all-apply 'vla-get-Area (list vlaobj)))
        (setq perim_raw (vl-catch-all-apply 'vla-get-Length (list vlaobj)))
        (if (and (numberp area_raw) (numberp perim_raw))
          (progn
            (vla-GetBoundingBox vlaobj 'minpt 'maxpt)
            (setq p1 (vlax-safearray->list minpt))
            (setq p2 (vlax-safearray->list maxpt))
            (setq len_m (* (abs (- (car p2) (car p1))) sc))
            (setq wid_m (* (abs (- (cadr p2) (cadr p1))) sc))
            (setq area_m2 (* area_raw sc sc))
            (setq perim_m (* perim_raw sc))
            (setq cnt (1+ cnt))
            (write-line
              (strcat
                cat ","
                lay ","
                mark_prefix (itoa cnt) ","
                "1,"
                (rtos area_m2 2 3) ","
                (rtos perim_m 2 3) ","
                (rtos len_m 2 3) ","
                (rtos wid_m 2 3) ","
                (rtos ht 2 3) ","
                (rtos sc 2 4)
              )
              fp
            )
          )
        )
        (setq i (1+ i))
      )
      (close fp)
      (princ (strcat "\n[AutoBOQ] Exported " (itoa cnt) " items to: " fpath))
      (princ "\nUpload this CSV in AutoBOQ 2D (Tab 4) to generate the 9-Tab Modular Excel BOQ.")
    )
  )
  (princ)
)

(defun c:AB_FOOTING () (_ab-export-selection "FOOTING" 0.525))
(defun c:AB_COLUMN  () (_ab-export-selection "COLUMN"  3.300))
(defun c:AB_SLAB    () (_ab-export-selection "SLAB"    0.125))
(defun c:AB_ROOMS   () (_ab-export-selection "ROOM"    3.000))
(defun c:AB_ALL     () (_ab-export-selection nil       3.000))

(princ "\nAutoBOQ_SmartTools.lsp loaded! Commands: AB_FOOTING | AB_COLUMN | AB_SLAB | AB_ROOMS | AB_ALL\n")
(princ)
