import hashlib
import json
import os
from datetime import datetime
from pathlib import Path


# ─────────────────────────────────────────────
# 1. SINGLE BLOCK
# Each piece of evidence = one block
# ─────────────────────────────────────────────
def _create_block(index, file_name, file_hash, investigator, previous_block_id):
    """
    Build one block in the chain.
    block_id is a hash of ALL fields combined — so nothing can be changed silently.
    """
    timestamp = datetime.now().isoformat()

    # Combine everything into one string, then hash it
    raw = f"{index}{file_name}{file_hash}{investigator}{timestamp}{previous_block_id}"
    block_id = hashlib.sha256(raw.encode()).hexdigest()

    return {
        "index"            : index,
        "timestamp"        : timestamp,
        "file_name"        : file_name,
        "file_hash"        : file_hash,
        "investigator"     : investigator,
        "previous_block_id": previous_block_id,
        "block_id"         : block_id,
    }


# ─────────────────────────────────────────────
# 2. LOAD EXISTING CHAIN FROM DISK
# ─────────────────────────────────────────────
def _load_chain(chain_path):
    """Load the existing blockchain from a JSON file. Returns empty list if none."""
    chain_path = Path(chain_path)
    if not chain_path.exists():
        return []
    try:
        with open(chain_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


# ─────────────────────────────────────────────
# 3. SAVE CHAIN TO DISK
# ─────────────────────────────────────────────
def _save_chain(chain, chain_path):
    """Save the blockchain to a JSON file."""
    chain_path = Path(chain_path)
    chain_path.parent.mkdir(parents=True, exist_ok=True)
    with open(chain_path, "w", encoding="utf-8") as f:
        json.dump(chain, f, indent=2)


# ─────────────────────────────────────────────
# 4. ADD EVIDENCE TO THE CHAIN
# Call this every time a new image is added
# ─────────────────────────────────────────────
def log_evidence(file_name, file_hash, investigator, chain_path):
    """
    Add a new evidence entry to the blockchain.

    Usage:
        log_evidence(
            file_name   = "photo_car.jpg",
            file_hash   = "aaa111...",
            investigator= "Dr. Ahmed",
            chain_path  = "cases/case_1/blockchain.json"
        )
    """
    chain = _load_chain(chain_path)

    # Get the previous block's ID (or "0000" if this is the first block)
    if chain:
        previous_block_id = chain[-1]["block_id"]
    else:
        previous_block_id = "0" * 64  # genesis block

    index = len(chain)
    block = _create_block(index, file_name, file_hash, investigator, previous_block_id)
    chain.append(block)
    _save_chain(chain, chain_path)

    return block


# ─────────────────────────────────────────────
# 5. VERIFY THE ENTIRE CHAIN
# Check every block is unbroken
# ─────────────────────────────────────────────
def verify_chain(chain_path):
    """
    Verify the entire blockchain is intact and unmodified.

    Returns:
        (is_valid, results_list)

        is_valid     — True if everything is fine
        results_list — one entry per block with pass/fail details
    """
    chain = _load_chain(chain_path)

    if not chain:
        return True, [{"status": "EMPTY", "detail": "No images logged yet — chain is empty."}]

    results = []
    is_valid = True

    for i, block in enumerate(chain):
        # Re-compute what the block_id SHOULD be
        raw = (
            f"{block['index']}"
            f"{block['file_name']}"
            f"{block['file_hash']}"
            f"{block['investigator']}"
            f"{block['timestamp']}"
            f"{block['previous_block_id']}"
        )
        expected_id = hashlib.sha256(raw.encode()).hexdigest()

        # Check 1: block_id matches
        if block["block_id"] != expected_id:
            results.append({
                "index"    : i,
                "file_name": block["file_name"],
                "status"   : "TAMPERED",
                "detail"   : f"Block {i} hash mismatch — this evidence record was modified!"
            })
            is_valid = False
            continue

        # Check 2: previous_block_id links correctly
        if i > 0:
            expected_previous = chain[i - 1]["block_id"]
            if block["previous_block_id"] != expected_previous:
                results.append({
                    "index"    : i,
                    "file_name": block["file_name"],
                    "status"   : "BROKEN LINK",
                    "detail"   : f"Block {i} is disconnected from Block {i-1} — chain was broken!"
                })
                is_valid = False
                continue

        results.append({
            "index"    : i,
            "file_name": block["file_name"],
            "timestamp": block["timestamp"],
            "status"   : "VERIFIED",
            "detail"   : f"Block {i} is intact."
        })

    return is_valid, results


# ─────────────────────────────────────────────
# 6. GET FULL CHAIN SUMMARY
# For displaying in the UI or report
# ─────────────────────────────────────────────
def get_chain_summary(chain_path):
    """
    Returns a readable summary of the entire chain.
    Use this to display in the UI tab or PDF report.
    """
    chain = _load_chain(chain_path)
    is_valid, results = verify_chain(chain_path)

    return {
        "total_blocks" : len(chain),
        "chain_valid"  : is_valid,
        "chain"        : chain,
        "verification" : results,
    }


# ─────────────────────────────────────────────
# 7. GET CHAIN FILE PATH FOR A CASE
# Helper so every module uses the same path
# ─────────────────────────────────────────────
def get_chain_path(artifact_dir):
    """Returns the standard blockchain file path for a case."""
    return Path(artifact_dir) / "blockchain.json"