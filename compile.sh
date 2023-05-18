#!/usr/bin/env bash

set -o errexit
set -o nounset
set -o pipefail
set -o noclobber

if [[ $OSTYPE == "darwin"* ]]; then
  echo MacOS
  # Assuming both are installed in the applications folder
  # Required some steps to make it work for MacOS M1, mind the `find` call which could return more than one (shouldn't)
  # install_name_tool -change @__VIA_LIBRARY_PATH__/libreglo.dylib $(find /Applications -name "libreglo.dylib") /Applications/LibreOffice7.4_SDK/bin/idlc
  # install_name_tool -change @__VIA_LIBRARY_PATH__/libuno_sal.dylib.3 $(find /Applications -name "libuno_sal.dylib.3") /Applications/LibreOffice7.4_SDK/bin/idlc
  # install_name_tool -change @__VIA_LIBRARY_PATH__/libuno_salhelpergcc3.dylib.3 $(find /Applications -name "libuno_salhelpergcc3.dylib.3") /Applications/LibreOffice7.4_SDK/bin/idlc
  # codesign --force -s - $(find /Applications -name "idlc")
  export PATH=$PATH:/Applications/LibreOffice7.4_SDK/bin
  export PATH=$PATH:/Applications/LibreOffice.app/Contents/MacOS
else
  export PATH=$PATH:/usr/lib/libreoffice/sdk/bin
  export PATH=$PATH:/usr/lib/libreoffice/program
fi

# Setup build directories

rm -rf "${PWD}"/build

mkdir "${PWD}"/build
mkdir "${PWD}"/build/META-INF/

# Compile the binaries

echo "Calling idlc..."
idlc -I/usr/lib/libreoffice/sdk/idl -w -verbose "${PWD}"/idl/XTCPToRTD.idl

echo "Calling regmerge..."
regmerge -v "${PWD}"/build/XTCPToRTD.rdb UCR "${PWD}"/idl/XTCPToRTD.urd

rm "${PWD}"/idl/XTCPToRTD.urd

echo "Generating meta files..."
python3 "${PWD}"/src/generate_metainfo.py

cp -f "${PWD}"/src/TCPToRTD.py "${PWD}"/build/

echo "Package into oxt file..."
pushd "${PWD}"/build/
zip -r "${PWD}"/RTD-Extension.zip ./*
popd

mv "${PWD}"/build/RTD-Extension.zip "${PWD}"/RTD-Extension.oxt