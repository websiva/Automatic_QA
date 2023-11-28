cd E:\001_QA_Generator\png_images
DEL /Q *.*
del /s /q "E:\001_QA_Generator\png_images*.*"
for /d %%p in ("E:\001_QA_Generator\png_images*.*") do rmdir "%%p" /s /q
cd E:\001_QA_Generator\downloaded_files\downloaded_pdf
DEL /Q *.* 
cd E:\001_QA_Generator\snap_tool\input\prt
DEL /Q *.*
del /s /q "E:\001_QA_Generator\snap_tool\input\prt\*.*"
for /d %%p in ("E:\001_QA_Generator\snap_tool\input\prt\*.*") do rmdir "%%p" /s /q
cd E:\001_QA_Generator\snap_tool\input\sldprt
DEL /Q *.*
del /s /q "E:\001_QA_Generator\snap_tool\input\sldprt\*.*"
for /d %%p in ("E:\001_QA_Generator\snap_tool\input\sldprt\*.*") do rmdir "%%p" /s /q
cd E:\001_QA_Generator\snap_tool\input\stp_simp
DEL /Q *.*
del /s /q "E:\001_QA_Generator\snap_tool\input\stp_simp\*.*"
for /d %%p in ("E:\001_QA_Generator\snap_tool\input\stp_simp\*.*") do rmdir "%%p" /s /q
CD E:\001_QA_Generator
python "E:\001_QA_Generator\snap_tool\delete_excel_sheet.py"
exit