;;; ============================================================================
;;; AutoBOQ_SmartTools.lsp — Complete Single-Command (AD) AutoCAD BBS & BOQ Tool
;;; Features matching & extending @smartciviltools:
;;;   - Single command:  AD  (auto-generates its own DCL pop-up dialog!)
;;;   - 10 Building Parts: Footing, Column, Plinth Beam, Roof/Floor Beam,
;;;     Roof/Floor Slab, Staircase, Water Tank/Sump, Wall/Masonry, Room/Tiles,
;;;     and Auto-Detect All by Layer
;;;   - Built-in Structural, Rebar (IS 2502 / IS 13920) & Cost Rate inputs
;;;   - Draws native AutoCAD BBS & BOQ + Material Yield Table inside your DWG!
;;;   - Exports BOTH AutoBOQ_Takeoff.csv (for 12-Tab Excel BOQ) and a standalone
;;;     part-wise BBS & BOQ CSV report!
;;;
;;; HOW TO USE:
;;;   1. Type APPLOAD -> Select AutoBOQ_SmartTools.lsp -> Load
;;;   2. Type  AD  and press Enter!
;;; ============================================================================

(vl-load-com)

(defun _ad-dwg-dir (/ d)
  (setq d (getvar "DWGPREFIX"))
  (if (or (null d) (= d ""))
    (setq d (strcat (getenv "USERPROFILE") "\\Desktop\\"))
  )
  d
)

(defun _ad-csv-path ()
  (strcat (_ad-dwg-dir) "AutoBOQ_Takeoff.csv")
)

(defun _ad-write-header (fpath / fp)
  (setq fp (open fpath "w"))
  (write-line "Part_Category,Layer,Mark,Nos,Area_m2,Perimeter_m,Length_m,Width_m,Depth_or_Height_m,Unit_Scale" fp)
  (close fp)
)

(defun _ad-unit-scale (/ ins)
  (setq ins (getvar "INSUNITS"))
  (cond
    ((= ins 4) 0.001)   ; mm
    ((= ins 5) 0.01)    ; cm
    ((= ins 6) 1.0)     ; m
    ((= ins 1) 0.0254)  ; inches
    ((= ins 2) 0.3048)  ; feet
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
    ((wcmatch u "*STAIR*") "STAIRCASE")
    ((or (wcmatch u "*TANK*") (wcmatch u "*SUMP*")) "WATER_TANK")
    ((or (wcmatch u "*WALL*") (wcmatch u "*MAS*") (wcmatch u "*BRICK*")) "WALL")
    (T "ROOM")
  )
)

(defun _ad-create-dcl (/ dcl_file fp)
  (setq dcl_file (vl-filename-mktemp "AutoBOQ_AD_v2.dcl"))
  (setq fp (open dcl_file "w"))
  (foreach ln
    '(
      "ad_dialog : dialog {"
      "  label = \"AutoBOQ 2D — Smart Structural BBS & BOQ Tool (Command: AD)\";"
      "  : row {"
      "    : boxed_radio_column {"
      "      label = \"Step 1: Choose Building Part\";"
      "      key = \"part_grp\";"
      "      : radio_button { key = \"rb_ftg\";   label = \"1. Footing (Isolated / Combined)\"; value = \"1\"; }"
      "      : radio_button { key = \"rb_col\";   label = \"2. Column (RCC + IS 13920 Ties)\"; }"
      "      : radio_button { key = \"rb_pbeam\"; label = \"3. Plinth Beam / Tie Beam\"; }"
      "      : radio_button { key = \"rb_beam\";  label = \"4. Roof / Floor Beam\"; }"
      "      : radio_button { key = \"rb_slab\";  label = \"5. Roof / Floor Slab\"; }"
      "      : radio_button { key = \"rb_stair\"; label = \"6. Staircase (Waist Slab & Steps)\"; }"
      "      : radio_button { key = \"rb_tank\";  label = \"7. Water Tank / Underground Sump\"; }"
      "      : radio_button { key = \"rb_wall\";  label = \"8. Wall / Masonry (230mm / 115mm)\"; }"
      "      : radio_button { key = \"rb_room\";  label = \"9. Room (Tiles, Skirting & Paint)\"; }"
      "      : radio_button { key = \"rb_all\";   label = \"10. Auto-Detect All Parts by Layer\"; }"
      "    }"
      "    : column {"
      "      : boxed_column {"
      "        label = \"Step 2A: Geometry & Floors\";"
      "        : edit_box { key = \"eb_prefix\"; label = \"Mark Prefix (F,C,PB,B,S,ST,WT,W,R):\"; edit_width = 10; value = \"F\"; }"
      "        : edit_box { key = \"eb_ht\";     label = \"Depth / Height / Slab Thk (m):\";     edit_width = 10; value = \"0.525\"; }"
      "        : edit_box { key = \"eb_thk\";    label = \"Wall / Beam Width (m, if line):\";    edit_width = 10; value = \"0.230\"; }"
      "        : edit_box { key = \"eb_nos\";    label = \"Identical Floors / Multiplier:\";     edit_width = 10; value = \"1\"; }"
      "      }"
      "      : boxed_column {"
      "        label = \"Step 2B: Rebar (IS 2502) & Rates (INR)\";"
      "        : edit_box { key = \"eb_mdia\";   label = \"Main Bar Dia (mm):\";                 edit_width = 10; value = \"12\"; }"
      "        : edit_box { key = \"eb_mbar\";   label = \"Main Bars/Mem OR Spacing (mm):\";     edit_width = 10; value = \"150\"; }"
      "        : edit_box { key = \"eb_tdia\";   label = \"Tie / Dist Bar Dia (mm):\";           edit_width = 10; value = \"8\"; }"
      "        : edit_box { key = \"eb_tspc\";   label = \"Tie / Dist Spacing (mm):\";           edit_width = 10; value = \"150\"; }"
      "        : edit_box { key = \"eb_rcnc\";   label = \"Concrete M25 Rate (Rs/Cum):\";        edit_width = 10; value = \"7600\"; }"
      "        : edit_box { key = \"eb_rstl\";   label = \"Steel Fe-500D Rate (Rs/kg):\";        edit_width = 10; value = \"82\"; }"
      "        : edit_box { key = \"eb_rfrm\";   label = \"Shuttering Rate (Rs/Sqm):\";          edit_width = 10; value = \"620\"; }"
      "      }"
      "    }"
      "  }"
      "  : row {"
      "    : toggle { key = \"tg_table\"; label = \"Draw BBS & BOQ Table inside AutoCAD\"; value = \"1\"; }"
      "    : toggle { key = \"tg_clear\"; label = \"Clear old CSV (start fresh)\"; value = \"0\"; }"
      "  }"
      "  : row {"
      "    : button { key = \"accept\"; label = \"Step 3: Select Objects in Drawing >>\"; is_default = true; width = 34; }"
      "    : button { key = \"cancel\"; label = \"Cancel\"; is_cancel = true; width = 14; }"
      "  }"
      "}"
    )
    (write-line ln fp)
  )
  (close fp)
  dcl_file
)

(defun _ad-draw-acad-table (pt title rows_list / doc ms tbl nrows ncols r c row_data th)
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
  (setq ms  (if (= (getvar "CVPORT") 1) (vla-get-PaperSpace doc) (vla-get-ModelSpace doc)))
  (setq nrows (+ 1 (length rows_list)))
  (setq ncols 4)
  (setq th (* (max 0.15 (/ 0.25 (_ad-unit-scale))) 1.0)) ; scale text height cleanly to drawing units
  (setq tbl (vla-AddTable ms (vlax-3d-point pt) nrows ncols (* th 2.6) (* th 18.0)))
  (vla-put-RegenerateTableSuppressed tbl :vlax-true)
  (vla-SetText tbl 0 0 title)
  (vla-SetTextHeight tbl 2 (* th 1.30)) ; acTitleRow
  (vla-SetTextHeight tbl 4 (* th 1.05)) ; acHeaderRow
  (vla-SetTextHeight tbl 1 (* th 0.95)) ; acDataRow
  (vla-SetColumnWidth tbl 0 (* th 26.0))
  (vla-SetColumnWidth tbl 1 (* th 13.0))
  (vla-SetColumnWidth tbl 2 (* th 12.0))
  (vla-SetColumnWidth tbl 3 (* th 16.0))
  (setq r 1)
  (foreach row_data rows_list
    (setq c 0)
    (foreach val row_data
      (vla-SetText tbl r c val)
      (setq c (1+ c))
    )
    (setq r (1+ r))
  )
  (vla-put-RegenerateTableSuppressed tbl :vlax-false)
  tbl
)

(defun _ad-extract-and-build (cat prefix ht wid_fallback mult mdia mbar tdia tspc rcnc rstl rfrm draw_tbl clear_csv
                              / sc ss i ent vlaobj objname lay row_cat
                                area_raw perim_raw minpt maxpt p1 p2
                                area_m2 perim_m len_m wid_m fpath fp part_csv pfp
                                cnt tot_area tot_perim tot_len tot_vol tot_shut
                                steel_kg cement_bags sand_cum agg_cum
                                cost_conc cost_stl cost_frm cost_tot
                                unit_wt_m unit_wt_t n_main n_tie cut_m cut_t
                                ins_pt tbl_rows)
  (setq sc (_ad-unit-scale))
  (princ (strcat "\n[AutoBOQ AD] Selected Part: " cat " | Scale to meters: " (rtos sc 2 4)))
  (princ "\n[AutoBOQ AD] Step 3: Select polylines or lines in AutoCAD, then press ENTER: ")

  (setq ss (ssget '((0 . "LWPOLYLINE,POLYLINE,LINE"))))
  (if (null ss)
    (progn (princ "\n[AutoBOQ AD] No objects selected.") nil)
    (progn
      (setq fpath (_ad-csv-path))
      (if (or (= clear_csv "1") (not (findfile fpath)))
        (_ad-write-header fpath)
      )
      (setq fp (open fpath "a"))
      (setq cnt 0 i 0 tot_area 0.0 tot_perim 0.0 tot_len 0.0 tot_vol 0.0 tot_shut 0.0)
      (setq steel_kg 0.0)
      (setq unit_wt_m (/ (* mdia mdia) 162.2))
      (setq unit_wt_t (/ (* tdia tdia) 162.2))

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
              (progn
                (setq len_m perim_m)
                (setq wid_m wid_fallback)
                (setq area_m2 (* len_m wid_m))
              )
            )
            (setq cnt (1+ cnt))
            (setq tot_area  (+ tot_area  (* area_m2 mult)))
            (setq tot_perim (+ tot_perim (* perim_m mult)))
            (setq tot_len   (+ tot_len   (* len_m mult)))
            (setq tot_vol   (+ tot_vol   (* area_m2 ht mult)))

            ;; Part-specific Shuttering & IS 2502 Steel Calculation
            (cond
              ((= row_cat "FOOTING")
                (setq tot_shut (+ tot_shut (* perim_m ht mult)))
                (setq n_main (+ 1 (fix (/ (max 0.0 (- wid_m 0.1)) (max 0.05 (/ mbar 1000.0))))))
                (setq cut_m (+ (- len_m 0.1) (* 2.0 (max 0.0 (- ht 0.1))) (- (* 0.004 mdia))))
                (setq n_tie (+ 1 (fix (/ (max 0.0 (- len_m 0.1)) (max 0.05 (/ mbar 1000.0))))))
                (setq cut_t (+ (- wid_m 0.1) (* 2.0 (max 0.0 (- ht 0.1))) (- (* 0.004 mdia))))
                (setq steel_kg (+ steel_kg (* mult (+ (* n_main cut_m unit_wt_m) (* n_tie cut_t unit_wt_m)))))
              )
              ((= row_cat "COLUMN")
                (setq tot_shut (+ tot_shut (* perim_m ht mult)))
                (setq n_main (max 4 (fix mbar)))
                (setq cut_m (+ ht (* 0.040 mdia)))
                (setq n_tie (+ 1 (fix (/ ht (max 0.05 (/ tspc 1000.0))))))
                (setq cut_t (+ (* 2.0 (+ (max 0.1 (- len_m 0.08)) (max 0.1 (- wid_m 0.08)))) (* 2.0 (max 0.075 (* 0.010 tdia))) (- (* 0.012 tdia))))
                (setq steel_kg (+ steel_kg (* mult (+ (* n_main cut_m unit_wt_m) (* n_tie cut_t unit_wt_t)))))
              )
              ((or (= row_cat "BEAM") (= row_cat "PLINTH_BEAM"))
                (setq tot_shut (+ tot_shut (* (+ (* 2.0 ht) wid_m) len_m mult)))
                (setq n_main (max 4 (fix mbar)))
                (setq cut_m (+ len_m (* 0.100 mdia) (- (* 0.004 mdia))))
                (setq n_tie (+ 1 (fix (/ len_m (max 0.05 (/ tspc 1000.0))))))
                (setq cut_t (+ (* 2.0 (+ (max 0.1 (- wid_m 0.05)) (max 0.1 (- ht 0.05)))) (* 2.0 (max 0.075 (* 0.010 tdia))) (- (* 0.012 tdia))))
                (setq steel_kg (+ steel_kg (* mult (+ (* n_main cut_m unit_wt_m) (* n_tie cut_t unit_wt_t)))))
              )
              ((or (= row_cat "SLAB") (= row_cat "STAIRCASE") (= row_cat "WATER_TANK"))
                (setq tot_shut (+ tot_shut (* (+ area_m2 (* perim_m ht)) mult)))
                ;; Slab main + distribution mesh steel (IS 2502)
                (setq n_main (+ 1 (fix (/ (max 0.1 (- wid_m 0.04)) (max 0.05 (/ mbar 1000.0))))))
                (setq cut_m (+ len_m 0.30 (* 0.42 (max 0.05 (- ht 0.04 (* 0.001 mdia))))))
                (setq n_tie (+ 1 (fix (/ (max 0.1 (- len_m 0.04)) (max 0.05 (/ tspc 1000.0))))))
                (setq cut_t (+ wid_m 0.30))
                (setq steel_kg (+ steel_kg (* mult (+ (* n_main cut_m unit_wt_m) (* n_tie cut_t unit_wt_t)))))
              )
              (T
                ;; Wall / Room (no rebar)
                (setq tot_shut (+ tot_shut (* perim_m ht mult)))
              )
            )

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

      ;; Concrete Mix Yield (M25 nominal) & Cost
      (if (or (= cat "WALL") (= cat "ROOM"))
        (progn
          (setq cement_bags (* tot_vol 1.25))
          (setq sand_cum    (* tot_vol 0.27))
          (setq agg_cum     0.0)
          (setq cost_conc   (* tot_vol 6900.0))
          (setq cost_stl    0.0)
          (setq cost_frm    0.0)
        )
        (progn
          (setq cement_bags (* tot_vol 8.20))
          (setq sand_cum    (* tot_vol 0.44))
          (setq agg_cum     (* tot_vol 0.86))
          (setq cost_conc   (* tot_vol rcnc))
          (setq cost_stl    (* steel_kg rstl))
          (setq cost_frm    (* tot_shut rfrm))
        )
      )
      (setq cost_tot (+ cost_conc cost_stl cost_frm))

      ;; Write standalone part BOQ & BBS CSV report
      (setq part_csv (strcat (_ad-dwg-dir) "AutoBOQ_" cat "_Schedule.csv"))
      (setq pfp (open part_csv "w"))
      (write-line "Item Description,Calculated Quantity,Unit,Amount (INR)" pfp)
      (write-line (strcat "Part & Count (" prefix "), " (itoa cnt) " x " (itoa mult) " flrs,Nos,-") pfp)
      (write-line (strcat "Plan Area," (rtos tot_area 2 2) ",Sqm,-") pfp)
      (write-line (strcat "Concrete / Masonry Volume," (rtos tot_vol 2 2) ",Cum," (rtos cost_conc 2 2)) pfp)
      (write-line (strcat "Shuttering / Contact Area," (rtos tot_shut 2 2) ",Sqm," (rtos cost_frm 2 2)) pfp)
      (write-line (strcat "TMT Fe-500D Steel (BBS)," (rtos steel_kg 2 2) ",kg," (rtos cost_stl 2 2)) pfp)
      (write-line (strcat "Cement Required (50kg Bags)," (rtos cement_bags 2 1) ",Bags,Incl.") pfp)
      (write-line (strcat "M-Sand Required," (rtos sand_cum 2 2) ",Cum,Incl.") pfp)
      (write-line (strcat "20mm Coarse Aggregate," (rtos agg_cum 2 2) ",Cum,Incl.") pfp)
      (write-line (strcat "TOTAL ESTIMATED COST,," "INR," (rtos cost_tot 2 2)) pfp)
      (close pfp)

      ;; Draw Native AutoCAD BBS & BOQ Table inside DWG!
      (if (= draw_tbl "1")
        (progn
          (setq ins_pt (getpoint "\n[AutoBOQ AD] Step 4: Pick insertion point for BBS & BOQ Table in drawing: "))
          (if ins_pt
            (progn
              (setq tbl_rows
                (list
                  (list "ITEM DESCRIPTION" "QUANTITY" "UNIT" "AMOUNT (INR)")
                  (list (strcat "Selected Elements (" prefix ")") (strcat (itoa cnt) " x " (itoa mult)) "Nos" "-")
                  (list "Total Plan Area" (rtos tot_area 2 2) "Sq.m" "-")
                  (list "Total Concrete / Masonry Vol" (rtos tot_vol 2 2) "Cu.m" (rtos cost_conc 2 0))
                  (list "Formwork / Shuttering Area" (rtos tot_shut 2 2) "Sq.m" (rtos cost_frm 2 0))
                  (list (strcat "Reinforcement Steel (Y" (rtos mdia 2 0) "/Y" (rtos tdia 2 0) ")") (rtos steel_kg 2 2) "kg" (rtos cost_stl 2 0))
                  (list "Cement (50 kg Bags)" (rtos cement_bags 2 1) "Bags" "Incl.")
                  (list "M-Sand (Fine Aggregate)" (rtos sand_cum 2 2) "Cu.m" "Incl.")
                  (list "20 mm Coarse Aggregate" (rtos agg_cum 2 2) "Cu.m" "Incl.")
                  (list "TOTAL ESTIMATED COST" "" "INR" (rtos cost_tot 2 0))
                )
              )
              (_ad-draw-acad-table ins_pt (strcat "AUTOBOQ BBS & BOQ SUMMARY — " cat) tbl_rows)
            )
          )
        )
      )

      (alert
        (strcat
          "AutoBOQ Smart Takeoff Complete!\n\n"
          "Part: " cat "  (" (itoa cnt) " items x " (itoa mult) " flr)\n"
          "Volume: " (rtos tot_vol 2 2) " Cum  |  Shuttering: " (rtos tot_shut 2 2) " Sqm\n"
          "Steel (BBS): " (rtos steel_kg 2 2) " kg  |  Cement: " (rtos cement_bags 2 1) " Bags\n"
          "Estimated Cost: Rs. " (rtos cost_tot 2 0) "\n\n"
          "Saved Master CSV: " fpath "\n"
          "Saved Part CSV: " part_csv "\n\n"
          "Type AD again for the next building part, or upload AutoBOQ_Takeoff.csv in AutoBOQ 2D Tab 4!"
        )
      )
      T
    )
  )
)

;;; ============================================================================
;;; SINGLE MAIN COMMAND:  AD
;;; ============================================================================
(defun c:AD (/ dcl_file dcl_id status sel_cat sel_prefix sel_ht sel_thk sel_nos
               sel_mdia sel_mbar sel_tdia sel_tspc sel_rcnc sel_rstl sel_rfrm
               sel_tbl sel_clear)
  (setq dcl_file (_ad-create-dcl))
  (setq dcl_id (load_dialog dcl_file))
  (if (not (new_dialog "ad_dialog" dcl_id))
    (progn
      (alert "Error: Could not load AutoBOQ AD pop-up dialog.")
      (exit)
    )
  )

  (setq sel_cat "FOOTING")

  (action_tile "rb_ftg"   "(setq sel_cat \"FOOTING\")     (set_tile \"eb_prefix\" \"F\")  (set_tile \"eb_ht\" \"0.525\") (set_tile \"eb_thk\" \"1.800\") (set_tile \"eb_mdia\" \"12\") (set_tile \"eb_mbar\" \"150\") (set_tile \"eb_tdia\" \"10\") (set_tile \"eb_tspc\" \"150\")")
  (action_tile "rb_col"   "(setq sel_cat \"COLUMN\")      (set_tile \"eb_prefix\" \"C\")  (set_tile \"eb_ht\" \"3.300\") (set_tile \"eb_thk\" \"0.300\") (set_tile \"eb_mdia\" \"20\") (set_tile \"eb_mbar\" \"8\")   (set_tile \"eb_tdia\" \"8\")  (set_tile \"eb_tspc\" \"125\")")
  (action_tile "rb_pbeam" "(setq sel_cat \"PLINTH_BEAM\") (set_tile \"eb_prefix\" \"PB\") (set_tile \"eb_ht\" \"0.380\") (set_tile \"eb_thk\" \"0.230\") (set_tile \"eb_mdia\" \"16\") (set_tile \"eb_mbar\" \"6\")   (set_tile \"eb_tdia\" \"8\")  (set_tile \"eb_tspc\" \"150\")")
  (action_tile "rb_beam"  "(setq sel_cat \"BEAM\")        (set_tile \"eb_prefix\" \"B\")  (set_tile \"eb_ht\" \"0.450\") (set_tile \"eb_thk\" \"0.230\") (set_tile \"eb_mdia\" \"16\") (set_tile \"eb_mbar\" \"6\")   (set_tile \"eb_tdia\" \"8\")  (set_tile \"eb_tspc\" \"125\")")
  (action_tile "rb_slab"  "(setq sel_cat \"SLAB\")        (set_tile \"eb_prefix\" \"S\")  (set_tile \"eb_ht\" \"0.125\") (set_tile \"eb_thk\" \"1.000\") (set_tile \"eb_mdia\" \"10\") (set_tile \"eb_mbar\" \"150\") (set_tile \"eb_tdia\" \"8\")  (set_tile \"eb_tspc\" \"200\")")
  (action_tile "rb_stair" "(setq sel_cat \"STAIRCASE\")   (set_tile \"eb_prefix\" \"ST\") (set_tile \"eb_ht\" \"0.175\") (set_tile \"eb_thk\" \"1.200\") (set_tile \"eb_mdia\" \"12\") (set_tile \"eb_mbar\" \"150\") (set_tile \"eb_tdia\" \"8\")  (set_tile \"eb_tspc\" \"200\")")
  (action_tile "rb_tank"  "(setq sel_cat \"WATER_TANK\")  (set_tile \"eb_prefix\" \"WT\") (set_tile \"eb_ht\" \"2.000\") (set_tile \"eb_thk\" \"0.200\") (set_tile \"eb_mdia\" \"12\") (set_tile \"eb_mbar\" \"150\") (set_tile \"eb_tdia\" \"10\") (set_tile \"eb_tspc\" \"150\")")
  (action_tile "rb_wall"  "(setq sel_cat \"WALL\")        (set_tile \"eb_prefix\" \"W\")  (set_tile \"eb_ht\" \"3.000\") (set_tile \"eb_thk\" \"0.230\") (set_tile \"eb_mdia\" \"0\")  (set_tile \"eb_mbar\" \"0\")   (set_tile \"eb_tdia\" \"0\")  (set_tile \"eb_tspc\" \"0\")")
  (action_tile "rb_room"  "(setq sel_cat \"ROOM\")        (set_tile \"eb_prefix\" \"R\")  (set_tile \"eb_ht\" \"3.000\") (set_tile \"eb_thk\" \"1.000\") (set_tile \"eb_mdia\" \"0\")  (set_tile \"eb_mbar\" \"0\")   (set_tile \"eb_tdia\" \"0\")  (set_tile \"eb_tspc\" \"0\")")
  (action_tile "rb_all"   "(setq sel_cat \"ALL\")         (set_tile \"eb_prefix\" \"E\")  (set_tile \"eb_ht\" \"3.000\") (set_tile \"eb_thk\" \"0.230\") (set_tile \"eb_mdia\" \"12\") (set_tile \"eb_mbar\" \"150\") (set_tile \"eb_tdia\" \"8\")  (set_tile \"eb_tspc\" \"150\")")

  (action_tile "accept"
    "(setq sel_prefix (get_tile \"eb_prefix\")
           sel_ht     (atof (get_tile \"eb_ht\"))
           sel_thk    (atof (get_tile \"eb_thk\"))
           sel_nos    (max 1 (atoi (get_tile \"eb_nos\")))
           sel_mdia   (atof (get_tile \"eb_mdia\"))
           sel_mbar   (atof (get_tile \"eb_mbar\"))
           sel_tdia   (atof (get_tile \"eb_tdia\"))
           sel_tspc   (atof (get_tile \"eb_tspc\"))
           sel_rcnc   (atof (get_tile \"eb_rcnc\"))
           sel_rstl   (atof (get_tile \"eb_rstl\"))
           sel_rfrm   (atof (get_tile \"eb_rfrm\"))
           sel_tbl    (get_tile \"tg_table\")
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
      (_ad-extract-and-build
        sel_cat sel_prefix sel_ht sel_thk sel_nos
        sel_mdia sel_mbar sel_tdia sel_tspc sel_rcnc sel_rstl sel_rfrm
        sel_tbl sel_clear
      )
    )
    (princ "\n[AutoBOQ AD] Cancelled.")
  )
  (princ)
)

(princ "\n======================================================================\n")
(princ " AutoBOQ_SmartTools.lsp v2 loaded! Type  AD  for Pop-Up BBS & BOQ!\n")
(princ "======================================================================\n")
(princ)
