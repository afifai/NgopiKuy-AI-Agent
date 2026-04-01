import os

OWNER_ID = str(os.getenv("TELEGRAM_OWNER_ID", "")).strip()

def block_if_not_owner(context: dict):
    """
    Return dict:
    - allow: bool
    - message: str (optional)
    """

    chat_type = context.get("chat_type") or context.get("platform_chat_type")
    user_id = str(context.get("user_id") or context.get("sender_id"))

    if chat_type == "private":
        if OWNER_ID and user_id != OWNER_ID:
            return {
                "allow": False,
                "message": "Maaf, bot ini hanya melayani owner di chat pribadi. Silakan gunakan di grup dengan mention @ngopikuy_bot."
            }

    return {"allow": True}
