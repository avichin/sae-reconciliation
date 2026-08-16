# SAE Reconciliation

SAE reconciliation between vendor and EDC (Electronic Data Capture) systems for clinical trial patient data. Usually performed manually by data managers.


<!-- Add visual anchors or relevant status badges below -->
[![License: GNU GLP v3.0](https://shields.io)](LICENSE)

---

## Table of Contents
- [Purpose](#-purpose)
- [Features](#-features)
- [Prerequisites](#%EF%B8%8F-prerequisites)
- [Getting Started](#-getting-started)
- [Usage](#-usage)
- [Contributing](#-contributing)
- [License](#-license)

---

## Purpose
The purpose of this project is to minimize manual reconciliation of trial data by data managers. Reconciliations between EDC and vendor data occur at pre-determined frequencies (typically quarterly, but as often as monthly) and are comprised of two files: the raw EDC listing and the vendor listing. 

Done manually, the data manager will spend hours comparing each field of each SAE record in the EDC with the data in the vendor listing. With this project, data managers will immediately receive an output with all the matching and mismatching SAE records with changes and comments from the previous reconciliation marked. As with all clinical data, it is imperative that every output is reviewed carefully by the designated professional. The intent of the output is not to replace data management skills, but rather to organize raw data for efficient review during trial enrollment. Upon receiving the output, the data manager will then review the discrepancies and post queries in the EDC or escalate as needed. 

This project was developed by Adeline Chin is currently in use by a leading oncology CRO. 

---

## Features
* **Core capability A** - Short description of why this matters.
* **Core capability B** - Highlight speed, safety, or efficiency.
* **Tracked Changes** - When issues are identified, the reviewing data manager's comments will be carried through subsequent outputs until the issue is resolved. All new records are highlighted in yellow and all changed records are highlighted in orange. 

---

## Prerequisites
Specify the minimum environment requirements needed to run the software successfully:
* **Node.js** v18.0 or higher
* **Docker** v20.10+ (optional, for containerized environments)
* **API Key** from [Service Name](https://example.com)

---

## Getting Started

Follow these steps to set up the project locally for development and testing.

### 1. Clone the Repository
```bash
git clone https://github.com
cd repository-name
```

### 2. Install Dependencies
```bash
npm install
```

### 3. Environment Configuration
Create a `.env` file in the root directory and add your keys:
```env
API_KEY=your_secret_key_here
PORT=3000
```

### 4. Run the Application
```bash
npm run dev
```

---

## Usage

Provide clear examples or commands showing how to interact with the project:

```javascript
import { client } from 'project-title';

// Initialize the client
const instance = client.init({ key: process.env.API_KEY });

// Execute primary functionality
const result = await instance.doSomething();
console.log(result);
```

*Include a screenshot, terminal recording, or architecture diagram here if applicable to enhance clarity.*

---

## Contributing

Contributions are what make the open-source community an amazing place to learn, inspire, and create. 

1. Review our [Contributing Guidelines](CONTRIBUTING.md) to understand coding styles and issue tracking.
2. Fork the project.
3. Create your feature branch (`git checkout -b feature/AmazingFeature`).
4. Commit your changes (`git commit -m 'Add some AmazingFeature'`).
5. Push to the branch (`git push origin feature/AmazingFeature`).
6. Open a Pull Request.

---

## License

Distributed under the GNU GPL v3.0 License. See the [LICENSE](LICENSE) file for more information.
