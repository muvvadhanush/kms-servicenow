import os
from fastapi import FastAPI, HTTPException, Query
from pydantic import BaseModel
from typing import List, Optional
from dotenv import load_dotenv
import core_logic

# Load environment variables
load_dotenv()

app = FastAPI(title="SharePoint KB Assistant API")

# Global vector store instance
vector_store = None

@app.on_event("startup")
def startup_event():
    global vector_store
    store_path = os.getenv("KB_STORE_PATH", "./kb_store")
    vector_store = core_logic.load_vector_store(store_path)
    if not vector_store:
        print(f"Warning: Could not load vector store from {store_path}. Some endpoints will fail until initialized.")

class SearchResult(BaseModel):
    title: str
    filename: str
    url: str
    content: str
    similarity: float

class IncidentAnalysis(BaseModel):
    incident_number: str
    short_description: str
    relevant_articles: List[SearchResult]

class PostWorkNoteRequest(BaseModel):
    incident_sys_id: str
    article_index: int  # Index in the last search results, or we can just pass the whole article data

@app.get("/")
def read_root():
    return {"message": "SharePoint KB Assistant API is running"}

@app.get("/search", response_model=List[SearchResult])
def search(query: str, limit: int = 5):
    if not vector_store:
        raise HTTPException(status_code=503, detail="Vector store not initialized")
    
    results = core_logic.search_articles(vector_store, query, limit)
    return results

@app.get("/analyze/{incident_number}", response_model=IncidentAnalysis)
def analyze_incident(incident_number: str, limit: int = 3):
    if not vector_store:
        raise HTTPException(status_code=503, detail="Vector store not initialized")

    snow_url = os.getenv('SNOW_INSTANCE_URL')
    snow_user = os.getenv('SNOW_USERNAME')
    snow_pass = os.getenv('SNOW_PASSWORD')

    snow_client = core_logic.initialize_snow_client(snow_url, snow_user, snow_pass)
    if not snow_client:
        raise HTTPException(status_code=502, detail="Failed to connect to ServiceNow")

    incident = core_logic.get_incident_details(snow_client, incident_number)
    if not incident:
        raise HTTPException(status_code=404, detail=f"Incident {incident_number} not found")

    search_query = core_logic.create_search_query_from_incident(incident)
    results = core_logic.search_articles(vector_store, search_query, limit)

    return {
        "incident_number": incident['number'],
        "short_description": incident['short_description'],
        "relevant_articles": results
    }

@app.post("/post-note")
def post_note(incident_sys_id: str, article: SearchResult):
    snow_url = os.getenv('SNOW_INSTANCE_URL')
    snow_user = os.getenv('SNOW_USERNAME')
    snow_pass = os.getenv('SNOW_PASSWORD')

    snow_client = core_logic.initialize_snow_client(snow_url, snow_user, snow_pass)
    if not snow_client:
        raise HTTPException(status_code=502, detail="Failed to connect to ServiceNow")

    success = core_logic.post_work_note(snow_client, incident_sys_id, article.dict())
    if not success:
        raise HTTPException(status_code=500, detail="Failed to post work note")
    
    return {"status": "success", "message": "Work note posted"}

# Endpoint to trigger re-index from SharePoint
@app.post("/reindex")
def reindex():
    tenant_id = os.getenv('SP_TENANT_ID')
    client_id = os.getenv('SP_CLIENT_ID')
    client_secret = os.getenv('SP_CLIENT_SECRET')
    site_url = os.getenv('SP_SITE_URL')
    library_name = os.getenv('SP_LIBRARY_NAME')
    store_path = os.getenv("KB_STORE_PATH", "./kb_store")

    if not all([tenant_id, client_id, client_secret, site_url, library_name]):
        raise HTTPException(status_code=400, detail="Missing SharePoint credentials in environment")

    client = core_logic.initialize_sharepoint_client(tenant_id, client_id, client_secret)
    if not client:
         raise HTTPException(status_code=502, detail="Failed to connect to SharePoint")

    sp_docs = core_logic.get_sharepoint_documents(client, site_url, library_name)
    if not sp_docs:
        return {"status": "warning", "message": "No documents found to index"}

    documents = core_logic.create_kb_documents(sp_docs)
    global vector_store
    vector_store = core_logic.initialize_vector_store(documents, store_path)
    
    return {"status": "success", "message": f"Indexed {len(documents)} document chunks"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
