@echo on
@echo Cloning JCL
@echo off
git clone git://github.com/project-jedi/jcl.git jcl
cd .\jcl
git submodule update --init

cd ..\..
@echo on
@echo Cloning JVCL
@echo off
git clone git://github.com/project-jedi/jvcl.git jvcl
cd .\jvcl
git submodule update --init

pause