import os
import argparse
import logging
from dotenv import load_dotenv
import core_logic

# Load environment variables
load_dotenv()

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

def cmd_init(args):
    """Initialize the KB store by fetching from SharePoint."""
    tenant_id = args.tenant_id or os.getenv('SP_TENANT_ID')
    client_id = args.client_id or os.getenv('SP_CLIENT_ID')
    client_secret = args.client_secret or os.getenv('SP_CLIENT_SECRET')
    site_url = args.site_url or os.getenv('SP_SITE_URL')
    library_name = args.library_name or os.getenv('SP_LIBRARY_NAME')

    if not all([tenant_id, client_id, client_secret, site_url, library_name]):
        logger.error("Missing SharePoint credentials. Provide them via arguments or .env file.")
        return

    logger.info("Connecting to SharePoint...")
    client = core_logic.initialize_sharepoint_client(tenant_id, client_id, client_secret)
    if not client:
        return

    logger.info(f"Fetching documents from {site_url}/{library_name}...")
    sp_docs = core_logic.get_sharepoint_documents(client, site_url, library_name)
    if not sp_docs:
        logger.warning("No documents found in SharePoint.")
        return

    logger.info(f"Processing {len(sp_docs)} documents...")
    documents = core_logic.create_kb_documents(sp_docs)
    
    logger.info("Initializing vector store...")
    vector_store = core_logic.initialize_vector_store(documents, args.output)
    if vector_store:
        logger.info(f"Successfully initialized KB store at {args.output}")

def cmd_search(args):
    """Search for articles in the KB store."""
    vector_store = core_logic.load_vector_store(args.store)
    if not vector_store:
        logger.error(f"Could not load vector store from {args.store}. Run 'init' first.")
        return

    results = core_logic.search_articles(vector_store, args.query, args.limit)
    if not results:
        logger.info("No matching articles found.")
        return

    print(f"\nFound {len(results)} relevant articles:\n")
    for i, res in enumerate(results, 1):
        print(f"[{i}] {res['title']} (Score: {res['similarity']:.2f})")
        print(f"    URL: {res['url']}")
        if args.verbose:
            content = core_logic.clean_html_content(res['content'])
            print(f"    Content Preview: {content[:200]}...")
        print("-" * 40)

def cmd_analyze(args):
    """Analyze a ServiceNow incident and find relevant KB articles."""
    # ServiceNow credentials
    snow_url = args.snow_url or os.getenv('SNOW_INSTANCE_URL')
    snow_user = args.snow_user or os.getenv('SNOW_USERNAME')
    snow_pass = args.snow_pass or os.getenv('SNOW_PASSWORD')

    if not all([snow_url, snow_user, snow_pass]):
        logger.error("Missing ServiceNow credentials.")
        return

    logger.info(f"Connecting to ServiceNow: {snow_url}...")
    snow_client = core_logic.initialize_snow_client(snow_url, snow_user, snow_pass)
    if not snow_client:
        return

    logger.info(f"Fetching incident {args.incident}...")
    incident = core_logic.get_incident_details(snow_client, args.incident)
    if not incident:
        return

    print(f"\nIncident: {incident['number']}")
    print(f"Short Description: {incident['short_description']}")
    
    # Search
    vector_store = core_logic.load_vector_store(args.store)
    if not vector_store:
        logger.error(f"Could not load vector store. Run 'init' first.")
        return

    search_query = core_logic.create_search_query_from_incident(incident)
    results = core_logic.search_articles(vector_store, search_query, args.limit)

    if not results:
        logger.info("No matching KB articles found for this incident.")
        return

    print(f"\nTop {len(results)} Matching Articles:\n")
    for i, res in enumerate(results, 1):
        print(f"[{i}] {res['title']} (Score: {res['similarity']:.2f})")
        print(f"    URL: {res['url']}")
    
    if args.post and results:
        top_article = results[0]
        confirm = input(f"\nPost the contents of '{top_article['title']}' to {incident['number']}? (y/n): ")
        if confirm.lower() == 'y':
            success = core_logic.post_work_note(snow_client, incident['sys_id'], top_article)
            if success:
                logger.info("Successfully posted to ServiceNow work notes.")
            else:
                logger.error("Failed to post work note.")

def main():
    parser = argparse.ArgumentParser(description="SharePoint KB Assistant CLI")
    subparsers = parser.add_subparsers(dest="command", help="Command to run")

    # Init command
    init_parser = subparsers.add_parser("init", help="Initialize the KB store from SharePoint")
    init_parser.add_argument("--tenant-id", help="Azure Tenant ID")
    init_parser.add_argument("--client-id", help="Azure Client ID")
    init_parser.add_argument("--client-secret", help="Azure Client Secret")
    init_parser.add_argument("--site-url", help="SharePoint Site URL")
    init_parser.add_argument("--library-name", help="Document Library Name")
    init_parser.add_argument("--output", default="./kb_store", help="Where to store the vector database")

    # Search command
    search_parser = subparsers.add_parser("search", help="Search the KB store")
    search_parser.add_argument("query", help="Text to search for")
    search_parser.add_argument("--limit", type=int, default=5, help="Max results")
    search_parser.add_argument("--store", default="./kb_store", help="Path to vector store")
    search_parser.add_argument("--verbose", action="store_true", help="Show article content preview")

    # Analyze command
    analyze_parser = subparsers.add_parser("analyze", help="Analyze an incident from ServiceNow")
    analyze_parser.add_argument("incident", help="Incident number (e.g. INC0012345)")
    analyze_parser.add_argument("--store", default="./kb_store", help="Path to vector store")
    analyze_parser.add_argument("--limit", type=int, default=3, help="Max results")
    analyze_parser.add_argument("--post", action="store_true", help="Option to post top result to ServiceNow")
    analyze_parser.add_argument("--snow-url", help="ServiceNow Instance URL")
    analyze_parser.add_argument("--snow-user", help="ServiceNow Username")
    analyze_parser.add_argument("--snow-pass", help="ServiceNow Password")

    args = parser.parse_args()

    if args.command == "init":
        cmd_init(args)
    elif args.command == "search":
        cmd_search(args)
    elif args.command == "analyze":
        cmd_analyze(args)
    else:
        parser.print_help()

if __name__ == "__main__":
    main()
