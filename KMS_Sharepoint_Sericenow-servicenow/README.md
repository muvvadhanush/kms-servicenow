# SharePoint KB Assistant

A Streamlit application that integrates SharePoint Knowledge Base articles with ServiceNow Incidents using Retrieval-Augmented Generation (RAG).

## Features

- **SharePoint Integration**: Fetches documents (PDF, Text) from SharePoint document libraries using Microsoft Graph API.
- **ServiceNow Integration**: Connects to ServiceNow instances to fetch incident details and update work notes.
- **Intelligent Search**: Uses LangChain and Chroma vector store to find the most relevant KB articles based on incident descriptions.
- **Automated Work Notes**: Automatically formats and posts relevant KB article content as work notes in ServiceNow.

## Architecture

1.  **Ingestion**: Documents are pulled from SharePoint and split into chunks.
2.  **Embedding**: Text chunks are converted into vector embeddings using `all-MiniLM-L6-v2`.
3.  **Storage**: Embeddings are stored in a local Chroma vector database (`kb_store/`).
4.  **Retrieval**: When an incident is analyzed, its description is used to query the vector store.
5.  **Output**: The top relevant article is retrieved and can be posted to the ServiceNow incident.

## Setup Instructions

### Prerequisites

- Python 3.8+
- Microsoft Azure App Registration (for SharePoint/Graph API access)
- ServiceNow Instance and credentials

### Installation

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd KMS_Sharepoint_Sericenow-servicenow
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Configure Environment:
   Create a `.env` file in the root directory (see `.env.example` if available or match the keys in `app.py`).
   ```env
   SNOW_INSTANCE_URL=your_instance.service-now.com
   SNOW_USERNAME=your_username
   SNOW_PASSWORD=your_password
   ```

### Running the Application

```bash
streamlit run app.py
```

## Deployment on Render

To deploy this Streamlit app on Render:

1.  **Create a New Web Service**: Connect your GitHub repository.
2.  **Environment**: Choose `Python 3`.
3.  **Build Command**: `pip install -r requirements.txt`
4.  **Start Command**: `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
5.  **Environment Variables**: Add your `.env` variables (e.g., `SNOW_INSTANCE_URL`, etc.) in the Render dashboard.

## Usage

1. **Configure SharePoint**: Enter your Azure Tenant ID, Client ID, Client Secret, Site URL, and Library Name in the sidebar.
2. **Load KB**: Click "Connect & Load KB" to ingest articles into the vector store.
3. **Analyze Incident**: Go to the "Incident Analysis" tab and enter a ServiceNow incident number (e.g., `INC0012345`).
4. **Review & Post**: Inspect the found articles and click "Post to Work Notes" to update ServiceNow.

## Project Structure

- `app.py`: Main Streamlit application.
- `kb_store/`: Persistent Chroma vector database.
- `requirements.txt`: Python package dependencies.
- `.env`: (Optional) Local secret storage for ServiceNow credentials.
