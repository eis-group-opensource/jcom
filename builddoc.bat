@echo off
@echo clear docs directory...
deltree /y docs
mkdir docs

set javadoc="javadoc"
set opt=-d docs -sourcepath . -encoding SJIS -splitindex %1 %2 %3 %4 %5 %6 %7 %8 %9
%javadoc% %opt% @packages
