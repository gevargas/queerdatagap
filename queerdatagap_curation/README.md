# QueerDataGap metadata and provenance package

Open `metadata_provenance_atlas.ipynb` in Jupyter and run cells in order. Keep the folder structure intact. Python 3.10 or later is required.

```sh
python -m pip install -r requirements.txt
jupyter lab metadata_provenance_atlas.ipynb
```

- `input/` contains byte-identical copies of the seven supplied files and two GitHub notebooks verified against commit `94fba639e34048cfe6ba57bddf8cd2ef5c606e4f`. The source files were not modified. Downloaded research notebooks are inspected, not executed.
- `input_manifest.json` documents the supplied files and checksums. `repository_snapshot.json` records repository URLs, commit and verified blob identifiers.
- `data_catalogue.xlsx` documents the collection using the first classroom table, extended with technical fields. Edit the creator, access, modification and evidence columns; annotations for `input/` paths are loaded by the notebook. Generated data are documented as a delivery snapshot.
- `metadata_schema.xlsx` defines reusable metadata fields and their extraction methods, provenance interpretation and Atlas mapping. It extends the second classroom table. It is documentation, not executable configuration.
- `runs/<run-id>/` contains CSV derivatives, structural metadata, observed process logs, PROV-JSON, Atlas type definitions and entity payloads, and validation results. Each execution creates a new folder.

No audio files or original CSV datasets were supplied. The notebook supports CSV input and creates CSV derivatives from the actual workbooks and transcripts. For additional CSVs, place them in `input/tables/`; the discovery step includes them, although initial package checks cover only manifested files. Register new inputs in the manifests for integrity verification.

The original podcast URLs and actual creators, collection dates, transcription methods, access rights and editing rights require curation. File metadata and analytical assignments are not evidence of those rights or roles. Atlas import is optional and disabled by default; no live server was contacted. The metadata model assumes Apache Atlas v2 with DataSet and Process base types. Adapt authentication and review custom types for your deployment.

The repository folder did not contain two catalogue templates. The deliverable uses the two tables requested in the earlier data-catalogue exercise and adapts them to the repository workflow. It does not claim to reproduce tables found in that repository.

This package includes the supplied research documents for local use. Raw transcript text stays in local CSV outputs; the Atlas payload contains metadata and lineage, not transcript or PDF body text.

## Included example run

The completed run is `runs/20260915T172858_c692996a/`. All nine code cells were executed in a Jupyter kernel. Local validation checks passed for 27 artifacts and 13 processes. The two Excel workbooks contain 27 catalogue records and 42 schema fields. See `verification.json` for the checks performed.

The long book filename was shortened in the local input copy to `Resentir_lo_queer_en_America_Latina.pdf`; its original name and unchanged content checksum are retained in the manifest.
