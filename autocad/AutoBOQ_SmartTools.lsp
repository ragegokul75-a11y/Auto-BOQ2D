;;; ============================================================================
;;; AutoBOQ_SmartTools.lsp — Single-Command (AD) Pop-Up AutoCAD to Excel BOQ
;;;
;;; HOW TO USE IN AUTOCAD:
;;;   1. Type APPLOAD -> Select AutoBOQ_SmartTools.lsp -> Click Load
;;;   2. Type AD and press Enter!
;;;   3. In the pop-up screen, pick:
;;;        - Footing
;;;        - Column
;;;        - Plinth Beam
;;;        - Roof / Floor Beam
;;;        - Roof / Floor Slab
;;;        - Wall / Masonry (230mm / 115mm)
;;;        - Room / Flooring / Paint
;;;        - Auto-Detect All by Layer
;;;   4. Check/edit Depth/Height, Thickness, Mark Prefix & Floor Multiplier ->
;;;      Click "Select Objects in Drawing >>" -> Pick polylines/lines -> Enter!
;;;   5. Upload AutoBOQ_Takeoff.csv into AutoBOQ 2D (Streamlit Tab 4) for your
;;;      complete 9-Tab Modular Excel BOQ!
;;; ============================================================================

(vl-load-com)

(defun _ad-csv-path (/ dwgdir)
  (setq dwgdir (getvar "DWGPREFIX"))
  (if (or (null dwgdir) (= dwgdir ""))
    (setq dwgdir (strcat (getenv "USERPROFILE") "\\Desktop\\"))
  )
  (strcat dwgdir "AutoBOQ_Takeoff.csv")
)

(defun _ad-write-header (fpath / fp)
  (setq fp (open fpath "w"))
  (write-line "Part_Category,Layer,Mark,Nos,Area_m2,Perimeter_m,Length_m,Width_m,Depth_or_Height_m,Unit_Scale" fp)
  (close fp)
)

(defun _ad-unit-scale (/ ins)
  ;; Convert drawing units to meters (4=mm, 5=cm, 6=m, 1=in, 2=ft)
  (setq ins (getvar "INSUNITS"))
  (cond
    ((= ins 4) 0.001)
    ((= ins 5) 0.01)
    ((= ins 6) 1.0)
    ((= ins 1) 0.0254)
    ((= ins 2) 0.3048)
    (T 0.001)
  )
)

(defun _ad-classify-layer (lay / u)
  (setq u (strcase lay))
  (cond
    ((or (wcmatch u "*FOOT*") (wcmatch u "*FTG*") (wcmatch u "*FDN*")) "FOOTING")
    ((or (wcmatch u "*COL*") (wcmatch u "*PIL*")) "COLUMN")
    ((wcmatch u "*PLINTH*") "PLINTH_BEAM")
    ((or (wcmatch u "*BEAM*") (wcmatch u "*BM*")) "BEAM")
    ((or (wcmatch u "*SLAB*") (wcmatch u "*ROOF*")) "SLAB")
    ((or (wcmatch u "*WALL*") (wcmatch u "*MAS*") (wcmatch u "*BRICK*")) "WALL")
    (T "ROOM")
  )
)

(defun _ad-create-dcl (/ dcl_file fp)
  (setq dcl_file (vl-filename-mktemp "AutoBOQ_AD.dcl"))
  (setq fp (open dcl_file "w"))
  (foreach ln
    '(
      "ad_dialog : dialog {"
      "  label = \"AutoBOQ 2D — Smart Building Part Takeoff (Command: AD)\";"
      "  : boxed_radio_column {"
      "    label = \"Step 1: Choose Building Part\";"
      "    key = \"part_grp\";"
      "    : radio_button { key = \"rb_ftg\";   label = \"1. Footing (Isolated / Combined)\"; value = \"1\"; }"
      "    : radio_button { key = \"rb_col\";   label = \"2. Column (RCC Vertical)\"; }"
      "    : radio_button { key = \"rb_pbeam\"; label = \"3. Plinth Beam / Tie Beam\"; }"
      "    : radio_button { key = \"rb_beam\";  label = \"4. Roof / Floor Beam\"; }"
      "    : radio_button { key = \"rb_slab\";  label = \"5. Roof / Floor Slab\"; }"
      "    : radio_button { key = \"rb_wall\";  label = \"6. Wall / Masonry (230mm / 115mm)\"; }"
      "    : radio_button { key = \"rb_room\";  label = \"7. Room (Flooring, Tiles, Skirting & Paint)\"; }"
      "    : radio_button { key = \"rb_all\";   label = \"8. Auto-Detect All Parts by Layer Name\"; }"
      "  }"
      "  : boxed_column {"
      "    label = \"Step 2: Parameters (Auto-Filled — Edit if Needed)\";"
      "    : edit_box { key = \"eb_prefix\"; label = \"Mark Prefix (F, C, PB, B, S, W, R):\"; edit_width = 12; value = \"F\"; }"
      "    : edit_box { key = \"eb_ht\";     label = \"Depth / Height / Slab Thk (m):\";     edit_width = 12; value = \"0.525\"; }"
      "    : edit_box { key = \"eb_thk\";    label = \"Wall / Beam Width (m, if line):\";    edit_width = 12; value = \"0.230\"; }"
      "    : edit_box { key = \"eb_nos\";    label = \"Identical Floors / Multiplier:\";     edit_width = 12; value = \"1\"; }"
      "    : toggle   { key = \"tg_clear\";  label = \"Start fresh CSV (clear previous exports)\"; value = \"0\"; }"
      "  }"
      "  : row {"
      "    : button { key = \"accept\"; label = \"Step 3: Select Objects in Drawing >>\"; is_default = true; width = 32; }"
      "    : button { key = \"cancel\"; label = \"Cancel\"; is_cancel = true; width = 12; }"
      "  }"
      "}"
    )
    (write-line ln fp)
  )
  (close fp)
  dcl_file
)

(defun _ad-extract-entities (cat prefix ht wid_fallback mult clear_csv
                             / sc ss i ent vlaobj objname lay row_cat
                               area_raw perim_raw minpt maxpt p1 p2
                               area_m2 perim_m len_m wid_m fpath fp cnt tot_area tot_len)
  (setq sc (_ad-unit-scale))
  (princ (strcat "\n[AutoBOQ AD] Part: " cat " | Scale to meters: " (rtos sc 2 4)))
  (princ "\n[AutoBOQ AD] Step 3: Select polylines or lines for this part, then press ENTER: ")

  ;; Allow LWPOLYLINE, POLYLINE, and LINE (useful for wall/beam centerlines)
  (setq ss (ssget '((0 . "LWPOLYLINE,POLYLINE,LINE"))))
  (if (null ss)
    (progn
      (princ "\n[AutoBOQ AD] No objects selected.")
      nil
    )
    (progn
      (setq fpath (_ad-csv-path))
      (if (or (= clear_csv "1") (not (findfile fpath)))
        (_ad-write-header fpath)
      )
      (setq fp (open fpath "a"))
      (setq cnt 0 i 0 tot_area 0.0 tot_len 0.0)
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        (setq vlaobj (vlax-ename->vla-object ent))
        (setq objname (vla-get-ObjectName vlaobj))
        (setq lay (vla-get-Layer vlaobj))
        (setq row_cat (if (= cat "ALL") (_ad-classify-layer lay) cat))

        (setq perim_raw (vl-catch-all-apply 'vla-get-Length (list vlaobj)))
        (setq area_raw  (vl-catch-all-apply 'vla-get-Area   (list vlaobj)))

        (if (numberp perim_raw)
          (progn
            (vla-GetBoundingBox vlaobj 'minpt 'maxpt)
            (setq p1 (vlax-safearray->list minpt))
            (setq p2 (vlax-safearray->list maxpt))
            (setq perim_m (* perim_raw sc))
            (if (and (numberp area_raw) (> area_raw 1e-6) (not (= objname "AcDbLine")))
              (progn
                (setq area_m2 (* area_raw sc sc))
                (setq len_m (* (abs (- (car p2) (car p1))) sc))
                (setq wid_m (* (abs (- (cadr p2) (cadr p1))) sc))
              )
              ;; Open LINE or zero-area centerline (e.g. Wall or Plinth Beam centerline)
              (progn
                (setq len_m perim_m)
                (setq wid_m wid_fallback)
                (setq area_m2 (* len_m wid_m))
              )
            )
            (setq cnt (1+ cnt))
            (setq tot_area (+ tot_area (* area_m2 mult)))
            (setq tot_len  (+ tot_len  (* len_m mult)))
            (write-line
              (strcat
                row_cat ","
                lay ","
                prefix (itoa cnt) ","
                (itoa mult) ","
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
      (alert
        (strcat
          "AutoBOQ Export Complete!\n\n"
          "Part: " cat "\n"
          "Selected Items: " (itoa cnt) " (x " (itoa mult) " floor/mult)\n"
          "Total Area: " (rtos tot_area 2 2) " sq.m\n"
          "Total Length: " (rtos tot_len 2 2) " m\n\n"
          "Saved to:\n" fpath "\n\n"
          "Type AD again to add another building part, or upload this CSV in AutoBOQ 2D Tab 4!"
        )
      )
      T
    )
  )
)

;;; ============================================================================
;;; MAIN SINGLE COMMAND: AD
;;; ============================================================================
(defun c:AD (/ dcl_file dcl_id status sel_cat sel_prefix sel_ht sel_thk sel_nos sel_clear)
  (setq dcl_file (_ad-create-dcl))
  (setq dcl_id (load_dialog dcl_file))
  (if (not (new_dialog "ad_dialog" dcl_id))
    (progn
      (alert "Error: Could not load AutoBOQ DCL pop-up dialog.")
      (exit)
    )
  )

  ;; Defaults (Footing selected initially)
  (setq sel_cat "FOOTING")

  ;; Auto-fill step-by-step defaults whenever user picks a part
  (action_tile "rb_ftg"   "(setq sel_cat \"FOOTING\")     (set_tile \"eb_prefix\" \"F\")  (set_tile \"eb_ht\" \"0.525\") (set_tile \"eb_thk\" \"1.800\")")
  (action_tile "rb_col"   "(setq sel_cat \"COLUMN\")      (set_tile \"eb_prefix\" \"C\")  (set_tile \"eb_ht\" \"3.300\") (set_tile \"eb_thk\" \"0.300\")")
  (action_tile "rb_pbeam" "(setq sel_cat \"PLINTH_BEAM\") (set_tile \"eb_prefix\" \"PB\") (set_tile \"eb_ht\" \"0.380\") (set_tile \"eb_thk\" \"0.230\")")
  (action_tile "rb_beam"  "(setq sel_cat \"BEAM\")        (set_tile \"eb_prefix\" \"B\")  (set_tile \"eb_ht\" \"0.450\") (set_tile \"eb_thk\" \"0.230\")")
  (action_tile "rb_slab"  "(setq sel_cat \"SLAB\")        (set_tile \"eb_prefix\" \"S\")  (set_tile \"eb_ht\" \"0.125\") (set_tile \"eb_thk\" \"1.000\")")
  (action_tile "rb_wall"  "(setq sel_cat \"WALL\")        (set_tile \"eb_prefix\" \"W\")  (set_tile \"eb_ht\" \"3.000\") (set_tile \"eb_thk\" \"0.230\")")
  (action_tile "rb_room"  "(setq sel_cat \"ROOM\")        (set_tile \"eb_prefix\" \"R\")  (set_tile \"eb_ht\" \"3.000\") (set_tile \"eb_thk\" \"1.000\")")
  (action_tile "rb_all"   "(setq sel_cat \"ALL\")         (set_tile \"eb_prefix\" \"E\")  (set_tile \"eb_ht\" \"3.000\") (set_tile \"eb_thk\" \"0.230\")")

  (action_tile "accept"
    "(setq sel_prefix (get_tile \"eb_prefix\")
           sel_ht     (atof (get_tile \"eb_ht\"))
           sel_thk    (atof (get_tile \"eb_thk\"))
           sel_nos    (max 1 (atoi (get_tile \"eb_nos\")))
           sel_clear  (get_tile \"tg_clear\"))
     (done_dialog 1)"
  )
  (action_tile "cancel" "(done_dialog 0)")

  (setq status (start_dialog))
  (unload_dialog dcl_id)
  (if (findfile dcl_file) (vl-file-delete dcl_file))

  (if (= status 1)
    (progn
      (if (= sel_prefix "") (setq sel_prefix "E"))
      (if (<= sel_ht 0.0) (setq sel_ht 3.0))
      (if (<= sel_thk 0.0) (setq sel_thk 0.23))
      (_ad-extract-entities sel_cat sel_prefix sel_ht sel_thk sel_nos sel_clear)
    )
    (princ "\n[AutoBOQ AD] Cancelled.")
  )
  (princ)
)

(princ "\n==============================================================\n")
(princ " AutoBOQ_SmartTools.lsp loaded! Type  AD  to open the Pop-Up!\n")
(princ "==============================================================\n")
(princ)
