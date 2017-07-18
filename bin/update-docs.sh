#!/bin/bash
lein do clean, codox
rm -rf docs/
cp -r target/doc docs
git add docs/
git commit -m "updated documentation"
git push -u origin master