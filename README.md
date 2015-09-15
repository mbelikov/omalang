# omalang
Omalang (office manipulation language / oma's language)

## introduction
I. Omalang model

F01. base office docs algorithms (load, save, search for a text / [P], search for a table / [Table])

F02. presentation / handling of 2 models: technical one (e.g. ODF representation), logical one (to use Omalang algorithms with)
F03. the technical model should remain by the Omalang manipulations (incl. all non-significant artefacts, e.g. ole-objects, formatting etc.)
F04. possible demarcation corresponding to the MVC pattern: model - "technical model", view - "logical model", controller - Omalang

F05. a doc is either a vector or a matrix of [P]'s and [Table]'s
F06. [P] := paragraph (text) (tech. model: it can contain separate text pieces along with some another objects between these pieces, e.g. ole-objects)
F07. [Table] := table containing [P]'s

F08. ability to construct a [P] and a [Table] (considering only the logical model)
F09. ability to insert / to remove a [P] or a [Table] to / from a doc
F10. ability to filter values from a vector / matrix to another vector / matrix
F11. ability to compare / fuzzy-compare vectors and matrices
F12. ability to subtract matrices / vectors
F13. ability to intersect matrices / vectors
F14. ability to replace / insert a vector to a matrix
F15. ...

a maven plugin :)

II. Omalang DSL / execution runtime

	D01. document basic I/O ops:
		exelDoc = load('ex_file_name.xlsx'); # it is in the same time a matrix
		wordDoc = load('word_file_name.docx'); # it is in the same time a vector of [P]'s and [Table]'s
		save('ex_new_file_name.xlsx', exelDoc);
		save('word_file_name.docx', wordDoc);
		
	D02. [P], [Table]:
		para = 'Some text here...';
		table = ['cell 11' 'cell 12', para; 'cell 21', para, 'cell 23'; 'cell 31' 'c 32' 'c 33']; # e.g. commas are optional
		para2 = 'Some text here...';
		para == para2 # result: true
		
	D03. vector / matrix manipulations like by the Octave, some examples:
		- 
		
	D04. document extended ops
		- 
		- formatting / existing format applying