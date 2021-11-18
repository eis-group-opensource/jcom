@echo off
echo clear temporary directory...
deltree /y build_tmp
mkdir build_tmp

echo compile all sources...
set java_home=d:\app\java
set javac=javac
set opt=-g -encoding SJIS -d build_tmp %1 %2 %3 %4 %5 %6 %7 %8 %9

@echo on
%javac% %opt% jp\ne\so_net\ga2\no_ji\jcom\*.java
%javac% %opt% jp\ne\so_net\ga2\no_ji\jcom\excel8\*.java
@echo off


REM JDK1.1 style
cd build_tmp
javah -jni -d ../cpp jp.ne.so_net.ga2.no_ji.jcom.IUnknown
javah -jni -d ../cpp jp.ne.so_net.ga2.no_ji.jcom.IDispatch
javah -jni -d ../cpp jp.ne.so_net.ga2.no_ji.jcom.ITypeInfo
javah -jni -d ../cpp jp.ne.so_net.ga2.no_ji.jcom.IEnumVARIANT
cd ..

@echo compile native files...
set cflag=/I%java_home%\include /I%java_home%\include\win32 /c
mkdir build_tmp\cpp
cl %cflag% /Fobuild_tmp/cpp/IUnknown.obj     cpp/IUnknown.cpp
cl %cflag% /Fobuild_tmp/cpp/IDispatch.obj    cpp/IDispatch.cpp
cl %cflag% /Fobuild_tmp/cpp/ITypeInfo.obj    cpp/ITypeInfo.cpp
cl %cflag% /Fobuild_tmp/cpp/IEnumVARIANT.obj cpp/IEnumVARIANT.cpp
cl %cflag% /Fobuild_tmp/cpp/IPersist.obj     cpp/IPersist.cpp
cl %cflag% /Fobuild_tmp/cpp/Com.obj          cpp/Com.cpp
cl %cflag% /Fobuild_tmp/cpp/callCom.obj      cpp/callCom.cpp
cl %cflag% /Fobuild_tmp/cpp/VARIANT.obj      cpp/VARIANT.cpp
cl %cflag% /Fobuild_tmp/cpp/jstring.obj      cpp/jstring.cpp
cl %cflag% /Fobuild_tmp/cpp/guid.obj         cpp/guid.cpp
cl %cflag% /Fobuild_tmp/cpp/InvokeHelper.obj cpp/InvokeHelper.cpp

@echo make dll...
link /dll build_tmp/cpp/IUnknown.obj build_tmp/cpp/IDispatch.obj build_tmp/cpp/ITypeInfo.obj build_tmp/cpp/IEnumVARIANT.obj build_tmp/cpp/IPersist.obj build_tmp/cpp/Com.obj build_tmp/cpp/callCom.obj build_tmp/cpp/VARIANT.obj build_tmp/cpp/jstring.obj build_tmp/cpp/guid.obj build_tmp/cpp/InvokeHelper.obj /OUT:jcom.dll
copy jcom.dll samples

@echo make executable jar file...
cd build_tmp
jar c0Mf ..\jcom.jar jp ../manifest.mf
cd ..

GOTO END
@echo create samples
cd samples
javac -classpath ../jcom.jar;%CLASSPATH%;. *.java
cd ..
@echo build complete.
:END