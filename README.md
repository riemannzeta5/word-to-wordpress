# Word to WordPress

This is a tool I developed to automatically convert my Microsoft Word content (notably including equations) to be suitable for publishing on my WordPress site. For all .docx files in the specified input folder, the tool will produce a .tex and a .html file in the specified output folder, where the HTML can be copied and pasted directly into WordPress's editor. Check the comments in the code for a more detailed usage description.

This works by using Pandoc and Luca Trevisan's LaTeX2WP program.

(Note: there is currently a bug causing content in subfolders to not be converted properly, so right now the script only converts top-level Word documents in the input folder.)

Hopefully this helps you too!  If you want to contribute, be aware that I don't actively maintain this project, but I'd still be happy to work with you on whatever improvements could be made.
