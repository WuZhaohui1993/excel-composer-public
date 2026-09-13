# Excel Composer

A browser-based tool for parsing Excel/PDF sources, mapping fields, combining tables, previewing results, and exporting a new workbook.

## Local development

```bash
npm install
npm run dev
```

The web client runs on port 5173 and the local API on port 3001. Set `ACCESS_TOKEN` only through a local `.env` file; never commit it.

The public preparation copy excludes customer documents, project-specific spreadsheets and PDFs, generated workbooks, deployment hosts, SSH material, and local AI logs. The remaining fixed preset code will be generalized in a later review.
