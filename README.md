# Reading Comprehension Study

This repository contains materials related to a reading comprehension study. It provides the web interface, stimulus assets, and scripts used to manage participant data.

## Repository Structure
- `index.html` – Landing page for the study interface.
- `stimuli/` – Image stimuli presented to participants.
- `api/`
  - `backup.js` – Node script for backing up participant data.
  - `wiat.gs` – Google Apps Script for interacting with the study API.

## Data and Privacy
Participant data is stored separately using anonymized IDs. The provided scripts handle backups and data access without exposing personally identifiable information.

## Usage
Run the backup script with Node.js to synchronize participant data:

```bash
node api/backup.js
```

This repository is intended for research purposes. Please ensure compliance with relevant privacy and ethical guidelines when handling study data.
