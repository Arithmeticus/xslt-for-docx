# Changes to Open and Save Archive

## Goals for 2019 revision
- XSLT 2.0 > 3.0
- Rename variables and temporary attributes to be more meaningful, and less likely to conflict with existing nodes.
- Give clearer instructions on how to import safely, so that the templates behave correctly.
- Make archive content retrieval more sound, based upon reading indexes and not presuming the existence of certain .rels components.
- For Saxon HE users
   - Support Word and Excel files, to open and save at least the XML components.
   - If any archive component (e.g., binary files) cannot be opened or saved, return a message notifying the user that the component is being skipped.
- For Saxon PE/EE users
   - Support any kind of archive.
   - Support opening and saving any kind of component, including binary.
   - Support conversion of archive maps to XML documents, and a reversal of the process, from XML documents to archive. 
   - Support a more robust form of `arch:extract-map()` than what is specified in the expath-archive module.
- Provide more and better use examples, each in its own subdirectory, clearly separating the input, output, and master stylesheet.
- Rename project to more accurately reflect its scope.

## Changes introduced
- Required that all input `@href` or uris be resolved before passing them to the functions. This includes converting spaces to %20.
- `@jar-path` renamed `@_archive-path`, to signal the temporary nature of the attribute and to point to the general nature of archives.
- Wholly rewrote most of the code, and wrote completely new stylesheets for the examples.
- Wrote `tan:map-to-xml()` but did not try to do the reverse. That may be included in a future release, if there is need.  