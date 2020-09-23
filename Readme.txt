http://www.dlcs.fsnet.co.uk

TreeView and ListView controls are not exactly well
documented and few good examples are currently available
to learn from.  What I hope I have achieved in this
small demo, is to bind these controls in such a way
that the VB application mimics MS Explorer.

There is a Form consisting of 7 Subs and the General 
declaration, the controls used are:- 2 ImageLists,
1 TreeView, 1 FileListBox, 1 DirListBox, 1 DriveListBox,
1 PictureBox and 1 ListView. I've attempted to keep the
code as small as possible so as to aid the learning 
process.

There are two API Functions one that deals with
Extracting the associated icon handle, and the DrawIcon
Function which binds the handle reference to the image
property of the PictureBox.  The picture property
records the contents of the image property and this in
turn is added to the ImageList control.  You then assign
the ImageList to the ListView control - e.g.

	     ListView.Icons = Imagelist
	     ListView.SmallIcons = ImageList

There is plently of scope to develop this demo much
further for various directory orientated projects.
A good Ftp program would benefit from a drag and drop
Explorer type interface.

There's no functionality other than browsing through
the directory folders, adding cut, copy and paste is
fairly routine and there are plently of examples around
to help you add these facilities and others.

Hope you find this helpful   -   Kev Heywood (uk)