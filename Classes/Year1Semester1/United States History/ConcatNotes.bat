@echo OFF
mkdir _TEMPCONCAT
echo Created temp folder
cd ToConcat
echo Navigating to src folder
dir /B /O:D > ../_TEMPCONCAT/_FILESTOCONCAT.txt
echo Creating file list
FOR /f %%f IN (../_TEMPCONCAT/_FILESTOCONCAT.txt) DO echo Appending %%f to output & type %%f >> ../ConcatNotes.txt
echo Done appending files
cd ..
echo Navigating to parent folder
rmdir /S /Q _TEMPCONCAT
echo Removing temp folder
echo DONE
pause
