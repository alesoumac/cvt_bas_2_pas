# cvt_bas_2_pas
cvt_bas_2_pas is a very "troglodyte" program that ("troglodytely") converts VB6 project to Lazarus project.

This program takes the FRM files and create the corresponding PAS and LFM files, focusing only in Form design (components and controls, positions, etc).

I created an alias in my .bashrc file, to ensure that I can run cb2p.py from any directory:

	$ alias cb2p='python3 ~/bin/scripts/cb2p.py'

So, we can convert our VB6 files, going to the project source directory and running "cb2p" alias command:

	user@computer:~$ cd \~/project_vb6/src
	user@computer:\~/project_vb6/src$ cb2p \*.frm

cb2p.py was programmed in Python 3 on Linux, but I think it works in Windows too, with a little difference: maybe you can't use '*.frm' as a parameter, thus you'll need to pass all the 'frm' file names, one by one, at the prompt command.
