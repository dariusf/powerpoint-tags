#!/usr/bin/env bash

set -x

mkdir _build
cp test.pptm _build

cd _build
unzip test.pptm
rm test.pptm
mkdir customUI
cp ../customUI14.xml customUI

zip -r test1.pptm *
cd -
mv _build/test1.pptm .
