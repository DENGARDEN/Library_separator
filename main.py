# # TODO: generalization
#
# ## Adapted: https://www.geeksforgeeks.org/how-to-iterate-through-excel-rows-in-python/
# ## https://stackoverflow.com/questions/22613272/how-to-access-the-real-value-of-a-cell-using-the-openpyxl-module-for-python
# ## User interaction should be implemented
# ## https://www.nature.com/articles/nmeth.4104
# # Will be integrated into Indel searcher
#
# # import module
# import os
#
# import openpyxl
#
# # load excel with its path
# wrkbk = openpyxl.load_workbook("1st order sheet.xlsx", data_only=True)
# sh = wrkbk.active
#
# # Magic numbers
# num_a_mm = 22270
# num_a_pam = 10240
# num_a_guide = 520
# num_a_full = 15104
#
# num_t_mm = 21181
# num_t_pam = 7727
# num_t_guide = 380
# num_t_full = 14893
#
# idx_a_mm = num_a_mm
# idx_a_pam = idx_a_mm + num_a_pam
# idx_a_guide = idx_a_pam + num_a_guide
# idx_a_full = idx_a_guide + num_a_full
#
# idx_t_mm = num_t_mm
# idx_t_pam = idx_t_mm + num_t_pam
# idx_t_guide = idx_t_pam + num_t_guide
# idx_t_full = idx_t_guide + num_t_full
#
# # Parameters
# ref_col = 1
# # length_col = 3
#
# FWD_HOMOLOGY_LENGTH = 27
# GUIDE_LENGTH = 20
# POLY_T_LENGTH = 7
# BARCODE_LENGTH = 24
# PAM_TYPE = "Cas12a"
# PAM_LENGTH = 4 if PAM_TYPE == "Cas12a" else 3
# TARGET_LENGTH = 26
# REV_HOMOLOGY_LENGTH = 22
# WINDOW_SIZE = 5
# WEDGE_POS = 22
# PATH = "./References"
#
# # Preliminary calculations
# s_barcode = FWD_HOMOLOGY_LENGTH + GUIDE_LENGTH + POLY_T_LENGTH
# e_barcode = s_barcode + BARCODE_LENGTH
#
# # Starting from PAM; TTTV
# s_guide = e_barcode
# e_guide = s_guide + PAM_LENGTH + GUIDE_LENGTH
#
# if not os.path.exists(PATH):
#     os.makedirs(PATH)
#
# barcode_path = os.path.join(PATH, "Barcode.txt")
# target_path = os.path.join(PATH, "Target_region.txt")
# ref_path = os.path.join(PATH, "Reference_sequence.txt")
#
# with open(barcode_path, 'w') as barcode, \
#         open(target_path, 'w') as target, \
#         open(ref_path, 'w') as ref:
#     # iterate through excel and display data
#     for i in range(1, sh.max_row + 1):
#         # assert (FWD_HOMOLOGY_LENGTH+GUIDE_LENGTH+POLY_T_LENGTH +BARCODE_LENGTH+ PAM_LENGTH +TARGET_LENGTH + REV_HOMOLOGY_LENGTH ) \
#         # 	== int(sh.cell(row=i , column =length_col).value), "Parameters are not matched with the source data"
#         print("\n")
#         print("Row ", i, " data :")
#         cell_obj = sh.cell(row=i, column=ref_col)
#         print(cell_obj.value, end=" ")
#
#         # Parameters are dynamically determined by current status
#         # TODO: Remove magic numbers...
#         if i <= idx_a_mm:
#             # AsCas12f1 MM
#             FWD_HOMOLOGY_LENGTH = 19
#             GUIDE_LENGTH = 20
#             POLY_T_LENGTH = 7
#             BARCODE_LENGTH = 20
#             PAM_TYPE = "AsCas12f1"
#             PAM_LENGTH = 8  # context involved
#             TARGET_LENGTH = 24
#             REV_HOMOLOGY_LENGTH = 18
#             WINDOW_SIZE = 10
#             WEDGE_POS = 22
#
#         elif i <= idx_a_pam:
#             # AsCas12f1 PAM
#             FWD_HOMOLOGY_LENGTH = 19
#             GUIDE_LENGTH = 20
#             POLY_T_LENGTH = 7
#             BARCODE_LENGTH = 20
#             PAM_TYPE = "AsCas12f1"
#             PAM_LENGTH = 8  # context involved
#             TARGET_LENGTH = 24
#             REV_HOMOLOGY_LENGTH = 18
#             WINDOW_SIZE = 10
#             WEDGE_POS = 22
#
#         elif i <= idx_a_guide:
#             # AsCas12f1 guide
#             FWD_HOMOLOGY_LENGTH = 19
#             POLY_T_LENGTH = 7
#             BARCODE_LENGTH = 20
#             PAM_TYPE = "AsCas12f1"
#             PAM_LENGTH = 8  # context involved
#             TARGET_LENGTH = 30
#             REV_HOMOLOGY_LENGTH = 18
#             WINDOW_SIZE = 10
#             WEDGE_POS = 22
#             # Variable guide lengths
#             GUIDE_LENGTH = len(
#                 cell_obj.value) - FWD_HOMOLOGY_LENGTH - POLY_T_LENGTH - BARCODE_LENGTH - PAM_LENGTH - TARGET_LENGTH \
#                            - REV_HOMOLOGY_LENGTH
#
#         elif i <= idx_a_full:
#             # AsCas12f1 FM
#             FWD_HOMOLOGY_LENGTH = 19
#             GUIDE_LENGTH = 20
#             POLY_T_LENGTH = 7
#             BARCODE_LENGTH = 20
#             PAM_TYPE = "AsCas12f1"
#             PAM_LENGTH = 38  # context involved
#             TARGET_LENGTH = 77
#             REV_HOMOLOGY_LENGTH = 18
#             WINDOW_SIZE = 10
#             WEDGE_POS = 22
#
#         elif i <= idx_t_mm:
#             # TnpB MM
#             FWD_HOMOLOGY_LENGTH = 26
#             GUIDE_LENGTH = 20
#             POLY_T_LENGTH = 7
#             BARCODE_LENGTH = 20
#             PAM_TYPE = "TnpB"
#             PAM_LENGTH = 9  # context involved
#             TARGET_LENGTH = 24
#             REV_HOMOLOGY_LENGTH = 18
#             WINDOW_SIZE = 10
#             WEDGE_POS = 20
#
#         elif i <= idx_t_pam:
#             # TnpB PAM
#             FWD_HOMOLOGY_LENGTH = 26
#             GUIDE_LENGTH = 20
#             POLY_T_LENGTH = 7
#             BARCODE_LENGTH = 20
#             PAM_TYPE = "TnpB"
#             PAM_LENGTH = 9  # context involved
#             TARGET_LENGTH = 24
#             REV_HOMOLOGY_LENGTH = 18
#             WINDOW_SIZE = 10
#             WEDGE_POS = 20
#
#         elif i <= idx_t_guide:
#             # TnpB guide
#             FWD_HOMOLOGY_LENGTH = 26
#             POLY_T_LENGTH = 7
#             BARCODE_LENGTH = 20
#             PAM_TYPE = "TnpB"
#             PAM_LENGTH = 9  # context involved
#             TARGET_LENGTH = 30
#             REV_HOMOLOGY_LENGTH = 18
#             WINDOW_SIZE = 10
#             WEDGE_POS = 20
#             # Variable guide lengths
#             GUIDE_LENGTH = len(
#                 cell_obj.value) - FWD_HOMOLOGY_LENGTH - POLY_T_LENGTH - BARCODE_LENGTH - PAM_LENGTH - TARGET_LENGTH \
#                            - REV_HOMOLOGY_LENGTH
#
#         elif i <= idx_t_full:
#             # TnpB FM
#             FWD_HOMOLOGY_LENGTH = 26
#             GUIDE_LENGTH = 20
#             POLY_T_LENGTH = 7
#             BARCODE_LENGTH = 20
#             PAM_TYPE = "TnpB"
#             PAM_LENGTH = 35  # context involved
#             TARGET_LENGTH = 71
#             REV_HOMOLOGY_LENGTH = 18
#             WINDOW_SIZE = 10
#             WEDGE_POS = 20
#
#     # Preliminary calculations
#     s_barcode = FWD_HOMOLOGY_LENGTH + GUIDE_LENGTH + POLY_T_LENGTH
#     e_barcode = s_barcode + BARCODE_LENGTH
#
#     # Starting from PAM; TTTV
#     s_guide = e_barcode
#     e_guide = s_guide + PAM_LENGTH + GUIDE_LENGTH
#
#     # Save as files
#     barcode.write(f"{cell_obj.value[s_barcode: e_barcode]}\n")
#     target.write(f"{cell_obj.value[s_guide: s_guide + PAM_LENGTH + WEDGE_POS + WINDOW_SIZE]}\n")
#     ref.write(f"{cell_obj.value}\n")
