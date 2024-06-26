#!/usr/bin/env bash

set -ex

rm -rf _build
mkdir _build
cp test.pptm _build

cd _build
unzip test.pptm
rm test.pptm

mkdir customUI
cp ../customUI14.xml customUI/customUI14.xml
sed -i '' 's@\(</Relationships>\)@<Relationship Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="/customUI/customUI14.xml" Id="Rabef4033419041b0" />\1@' _rels/.rels

zip -r test1.pptm *
cd -