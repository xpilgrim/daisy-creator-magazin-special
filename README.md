daisy-creator-magazin-special
=============================

Create digital talking books of audio magazins for blind people in the daisy 2.02 standard


basics
------

Basis of the functionality are the filenames of existing audio-files (mp3).
Additionally here is used a so called counter file to setup the structure of the dtb.

The syntax of filenames must be as follows:

author_-_title.mp3

> nnnn is a 4-digit number, it is usesd for right sorting

> author describes the last name of the author

> title is to describe the content of the file


The syntax of the counter file must be as follows:

nn_author_-_title.mp3

> nn is a two digit for the order of the items

While processing, the audio files will be renamed 
to start with the number taken from the counter file.

