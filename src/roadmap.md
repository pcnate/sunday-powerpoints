## todo checklist

 1. ~~convert to typescript~~
 2. ~~add folder creation mode~~
 3. ~~add a server.ts that accepts env options as described below~~
 4. ~~add a web server to the server~~
 5. ~~add song selection for each week by month~~
 6. add semantic-release automation for releases
 7. convert to docker (if possible)
 8. add bible verse extraction API endpoint
 9. add bible verse update API endpoint

## Convert to TypeScript
Convert the entire project into TypeScript to ease coding complexity

## folder creation mode
 1. Add a command line option to create folders so the template file doesn't need to be updated
 2. Needs to then allow the other files to be created and copied in those same folders later without failing

## server.ts
**Note:** dotenv support for .env files was added to server.ts (missed in original requirements)
 1. Create a server.ts that is intended be run as a service. In later steps, docker will be used
 2. add a script to run server.ts in a development mode, not necessarily node dev vs prod
 2. accept the env options in the table below

| option | required | default value |
| -- | -- | -- |
| templateFile | yes | none |
| ext | no | pptx |
| templateDirectory | no | /input |
| outputDirectory | no | /output |
| rootPath | no | %OneDriveConsumer% |

## web server
 1. add a basic express webserver on port 80 and serve files out of src/webapp/* or lib/src/webapp/* when no running in development mode
 2. add boostrap css and js for use in the next steps, the express server should serve these from the node_modules in development and I'm not sure how they should be included in production/docker but that can be a problem for later
 2. in the web ui, show a list of the folders in outputDirectory
 3. in the web ui, show the last date modified of the templateFile
 4. add a button to run the folder creation script as a child process with a checkbox to enable write and print the results onscreen

## Release Automation with semantic-release

Set up automated releases using semantic-release for the project. The release process should:
- Run on pushes to the `main`, `master`, and `development` branches.
- For `main` and `master`, generate standard releases (e.g., `x.y.z`).
- For `development`, generate release candidates using the `x.x.x-rc.x` format.
- Do not publish any npm package (no public npm publish step).
- Ensure the publish config is set to private (see `package.json`'s `"publishConfig": { "access": "restricted" }`).
- Use plugins for changelog generation, git tagging, and GitHub releases as needed.
- Optionally, update the changelog and version in the repository.

This ensures a robust, automated release workflow without exposing the package to the public npm registry.

## song selection

### user story
 1. I wish to choose the songs for each sunday of a selected month. Each sunday can have any number of choruses and 3 songs. I wish to select from a searchable dropdown all songs in the library for song 1-3. I want to add and select each chorus above the 3 songs.
 2. When a week is saved, it should create a shortcut to that song in the weeks folder (create the folder if necessary) where the name should be `${ 'song' or 'chorus' }$ ${ n } - ${ targetFileName }`. Create a JSON file in that folder that has each songs name and an unpopulated field for the CCLI license that we will implement someday.

### structure
 1. select a month
 2. have 4-5 columns (on desktop) that have a list of songs for each week
 3. on mobile, a week selection should be included

## semantic-release
 1. generate releases from main, master or development branches
 2. do not publish npm package
 3. ensure if any publishing, that it is private
 4. development branch creates a release candidated version number systax 1.2.3-RC.1
 5. set a boolean output for the semantic-release job that flags whether a release was created
 6. set an output for the semantic-release job that contains the version number being created. this will be used for publishing to GHCR etc.

## dockerization
 1. determine if the windows-shortcuts dependency can run in linux/docker
 2. if #1 works, then add a dockerfile, otherwise abandon the rest of this section
 3. add a .npmignore file that ignores all source code not found in ./lib
 4. ensure TypeScript (tsc) can build into ./lib using fairly new ecmascript
 5. build the docker file with the intention of publishing as a private container to GitHub Container Registry
 6. The .env file must be excluded from the Docker image (add to .dockerignore in the future)
 7. CI, a job should be run after semantic-release completes and detects a new version was created to publish to GHCR and tag it with the semantic-release generated tag

## Extracting Bible Verse Text from Template PPTX

To support automation and display of the Bible verse from the template PowerPoint, add an API endpoint (e.g., `/api/template-slide-text?slide=6`) that extracts and returns the text from a specific slide (defaulting to slide #6) of the configured template PPTX file.

**Possible approach:**
- Use a Node.js library such as `pptx-parse`, `pptx2json`, or `officeparser` to read and parse the PPTX file. These libraries can extract text from slides by parsing the XML content inside the zipped PPTX archive.
- Locate the template file using the configured environment variables (see `TEMPLATE_FILE` and `TEMPLATE_DIRECTORY`).
- Extract all text content from the requested slide (slide #6 by default). If the slide contains multiple text boxes, concatenate or return them as an array.
- Return the extracted text as JSON from the API endpoint for use in the web UI or other automation.
- Consider error handling for missing files, missing slides, or parsing errors.

**Example response:**
```json
{
  "slide": 6,
  "text": "For God so loved the world... (John 3:16)"
}
```

**Optional enhancements:**
- Allow the endpoint to accept a slide number as a query parameter.
- Cache the result for performance if the template does not change often.
- Add a UI element to display the extracted verse in the webapp.

## Updating Bible Verse Text in Template PPTX

To allow automation or UI-driven updates of the Bible verse on slide #6 of the template PowerPoint, add an API endpoint (e.g., `/api/update-template-slide-text`) that updates the text content of a specific slide (defaulting to slide #6) in the configured template PPTX file.

**Possible approach:**
- Use a Node.js library such as `pptxgenjs`, `officegen`, or `pizzip` + `xml2js` to modify the XML content of the PPTX file. These libraries or tools can update text in a slide by manipulating the zipped XML structure.
- Locate the template file using the configured environment variables (see `TEMPLATE_FILE` and `TEMPLATE_DIRECTORY`).
- Accept the new verse text and (optionally) the slide number in the API request body or query.
- Update all text boxes on the specified slide, or target a specific text box if needed, replacing the old verse with the new one.
- Save the updated PPTX file, either overwriting the original or creating a backup/copy as needed.
- Return a success/failure response as JSON, including error details if the update fails.
- Consider error handling for missing files, missing slides, or XML/ZIP manipulation errors.

**Example request:**
```json
{
  "slide": 6,
  "text": "New Bible verse text here (John 3:17)"
}
```

**Example response:**
```json
{
  "success": true,
  "slide": 6
}
```

**Optional enhancements:**
- Allow the endpoint to accept a slide number and/or text box index as parameters.
- Add versioning or backup of the template file before overwriting.
- Add a UI element to edit and update the verse directly from the webapp.

