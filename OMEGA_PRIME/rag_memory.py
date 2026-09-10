import logging
import hashlib

try:
    import chromadb
    from chromadb.config import Settings
    CHROMA_AVAILABLE = True
except ImportError:
    CHROMA_AVAILABLE = False

class EpisodicMemory:
    def __init__(self, db_path="./chroma_db"):
        logging.info("Initializing Local RAG System (ChromaDB) for Episodic Memory...")
        self.db_path = db_path
        if CHROMA_AVAILABLE:
            try:
                self.client = chromadb.PersistentClient(path=self.db_path)
                self.collection = self.client.get_or_create_collection(name="omega_prime_memories")
                logging.info("ChromaDB Collection 'omega_prime_memories' ready.")
            except Exception as e:
                logging.error(f"Failed to initialize ChromaDB: {e}")
                self.client = None
                self.collection = None
        else:
            logging.warning("ChromaDB library not found. Operating in MOCK RAG Mode.")
            self.client = None
            self.collection = None
            
    def _generate_id(self, text):
        return hashlib.md5(text.encode()).hexdigest()

    def store_memory(self, memory_text, metadata=None):
        logging.info("Storing new data into Episodic Memory (RAG)...")
        if not metadata:
            metadata = {"source": "omega-prime-epoch"}
            
        if self.collection:
            doc_id = self._generate_id(memory_text)
            self.collection.upsert(
                documents=[memory_text],
                metadatas=[metadata],
                ids=[doc_id]
            )
            logging.info(f"Memory successfully embedded and stored. [ID: {doc_id[:8]}]")
        else:
            logging.info(f"[MOCK RAG] Memory Stored: {memory_text[:50]}...")
            
    def retrieve_memory(self, query_text, n_results=1):
        logging.info(f"Querying Episodic Memory for: '{query_text}'...")
        if self.collection:
            # Check if collection is empty
            if self.collection.count() == 0:
                logging.info("Episodic Memory is empty. No past experiences found.")
                return []
                
            results = self.collection.query(
                query_texts=[query_text],
                n_results=n_results
            )
            
            if results['documents'] and results['documents'][0]:
                logging.info(f"Memory Recall Successful! Found {len(results['documents'][0])} past experiences.")
                return results['documents'][0]
            else:
                logging.info("No relevant past experiences found.")
                return []
        else:
            logging.info("[MOCK RAG] Query Complete. Memory Recall: 'Simulated past experience regarding query.'")
            return ["[MOCK RAG] Retrieved Memory Context"]
