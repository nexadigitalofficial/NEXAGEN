import sys
import logging
from orchestrator import OmegaPrimeOrchestrator

logging.basicConfig(level=logging.INFO, format='%(asctime)s - CLOUD-ACTION-TRIGGER - %(levelname)s - %(message)s')

def trigger_action():
    logging.info("Cloud Action Triggered (e.g., via GitHub Actions Cron).")
    try:
        # Initialize the core engine
        engine = OmegaPrimeOrchestrator()
        
        # Run a single autonomous epoch
        engine.run_epoch()
        
        logging.info("Cloud Action Epoch completed successfully.")
        sys.exit(0)  # Success exit code for CI/CD
    except Exception as e:
        logging.error(f"FATAL ERROR during Cloud Action execution: {str(e)}")
        sys.exit(1)  # Failure exit code for CI/CD to alert administrators

if __name__ == "__main__":
    trigger_action()
