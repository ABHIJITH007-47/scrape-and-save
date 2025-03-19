import pandas as pd
from Bio import Entrez

def fetch_pubmed_data(query, max_results=10):
    """Fetches article data from PubMed based on a search query."""
    
    Entrez.email = "abhijithbasar66@gmail.com"  # Replace with your real email
    Entrez.sleep_between_tries = 1  # Helps prevent API rate limits

    # Search PubMed
    try:
        print(f"Searching PubMed for: {query} (max {max_results} results)")
        search_handle = Entrez.esearch(db="pubmed", term=query, retmax=max_results)
        search_results = Entrez.read(search_handle)
        search_handle.close()
    except Exception as e:
        print(f"❌ Error during search: {e}")
        return []

    pmid_list = search_results.get("IdList", [])
    if not pmid_list:
        print("⚠️ No results found for the query.")
        return []

    # Fetch article details
    try:
        print(f"Fetching details for {len(pmid_list)} articles...")
        fetch_handle = Entrez.efetch(db="pubmed", id=",".join(pmid_list), rettype="xml", retmode="text")
        records = Entrez.read(fetch_handle)
        fetch_handle.close()
    except Exception as e:
        print(f"❌ Error fetching data: {e}")
        return []

    articles = []
    for article in records.get("PubmedArticle", []):
        title = article["MedlineCitation"]["Article"].get("ArticleTitle", "No title available")
        authors = [
            f"{author.get('LastName', '')} {author.get('ForeName', '')}".strip()
            for author in article["MedlineCitation"]["Article"].get("AuthorList", [])
        ]
        abstract = article["MedlineCitation"]["Article"].get("Abstract", {}).get("AbstractText", ["No abstract available"])[0]

        # Extract Publication Date
        pub_date = article["MedlineCitation"].get("Article", {}).get("Journal", {}).get("JournalIssue", {}).get("PubDate", {})
        year = pub_date.get("Year", "Unknown Year")
        month = pub_date.get("Month", "")
        day = pub_date.get("Day", "")
        full_date = f"{year}-{month}-{day}".strip("-")  # Format as YYYY-MM-DD

        articles.append({
            "Title": title,
            "Authors": ", ".join(authors) if authors else "No authors listed",
            "Abstract": abstract,
            "Publication Date": full_date
        })

    print(f"✅ Successfully fetched {len(articles)} articles!")
    return articles

def save_to_excel(articles, filename="pubmed_results_new.xlsx"):
    """Saves fetched articles to an Excel file."""
    
    if not articles:
        print("⚠️ No articles to save.")
        return
    
    df = pd.DataFrame(articles)  # Convert to Pandas DataFrame
    df.to_excel(filename, index=False)  # Save to Excel
    print(f"✅ Data successfully saved to: {filename}")

if __name__ == "__main__":
    query = "paracetamol"  # Change this to your search term
    articles = fetch_pubmed_data(query)
    save_to_excel(articles)
