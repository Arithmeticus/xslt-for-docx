# xslt-for-docx 2019

## Goals for 2019 revision
- XSLT 2.0 > 3.0
- Rename variables and temporary attributes to be more meaningful, and less likely to conflict with existing nodes.
- Give clearer instructions on how to import safely, so that the templates behave correctly.
- Make archive contents retrieval more sound, based upon reading indexes and not presuming the existence of certain components.
- For Saxon HE users
   - Support Word and Excel files, to open and save at least the XML components.
   - If any archive component (e.g., binary files) cannot be opened or saved, return a message notifying the user that the component is being skipped.
- For Saxon PE/EE users
   - Support opening and saving all components, including binary files.
   - Support any kind of archive.
   - Support conversion of archive maps to XML documents, and a reversalo of the process, so that any altered XML documents can be repackaged in an archive. 
   - Support a more robust form of arch:extract-map() than what is specified in the expath-archive module.
- Provide more and better use examples, each in its own subdirectory:
   - Unpack and repack components 
   - Scrape into plain text
   - Use real regular expressions to change a document 
   - Generate form letters
   - Anonymize comments

## Changes introduced
- All input @href or uris must be resolved before reaching the functions. This includes converting spaces to %20.
- @jar-path renamed @_jar-path, to signal the temporary nature of the attribute.


## Notes

Saxon HE does not support extended functions, such as reading and writing binary files, or getting the contents of an archive. One must know how a particular archive is designed, to navigate its contents. Currently only docx and xslx architecture is supported, but the library could be extended to other file types.

With advanced functions (Saxon PE/EE), archives within archives are retrieved fine as maps, but once converted to XML documents and saved as a new archive, that new archive will contain only other files, not archives. This feature may be supported in a future release.  