"""
send_reasons_whatsapp.py
------------------------
Envoie automatiquement une liste de "raisons pour lesquelles je t'aime"
ligne par ligne sur WhatsApp Web, via pyautogui.

UTILISATION
===========
1. Installe pyautogui :
        pip install pyautogui
2. Ouvre WhatsApp Web (https://web.whatsapp.com/) dans ton navigateur.
3. Clique sur le chat de ta copine pour qu'il soit ACTIF
   (le curseur doit etre dans la zone "Tapez un message").
4. Lance le script :
        python send_reasons_whatsapp.py
5. Tu auras 7 secondes pour revenir sur la fenetre WhatsApp Web.
   Ne touche plus a la souris ni au clavier ensuite.

ARRET D'URGENCE
===============
Bouge la souris dans un coin de l'ecran -> pyautogui fait un FailSafe
et arrete tout immediatement.
Sinon : Ctrl + C dans le terminal.

CONFIG
======
- DELAY_BETWEEN_MESSAGES : secondes entre chaque message
- START_DELAY           : secondes avant le debut (le temps de revenir
                          sur WhatsApp)
- START_INDEX           : index (1-based) auquel reprendre si tu veux
                          relancer apres une coupure
"""

import time
import sys

try:
    import pyautogui
    import pyperclip
except ImportError:
    print("[ERREUR] Modules manquants. Installe-les avec :")
    print("    pip install pyautogui pyperclip")
    sys.exit(1)


# ============================================================
# CONFIGURATION
# ============================================================
DELAY_BETWEEN_MESSAGES = 3.0   # secondes entre 2 messages (3s = safe contre les blocages WhatsApp)
START_DELAY = 7                # secondes avant de commencer
START_INDEX = 1                # reprendre a partir de cette raison (1 = depuis le debut)

# Securite : si la souris va dans un coin, pyautogui s'arrete
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.05


# ============================================================
# LA LISTE
# ============================================================
REASONS = [
    "your smile",
    "your laugh",
    "your eyes",
    "your personality",
    "your hair",
    "your humour",
    "your voice",
    "your style",
    "your name",
    "the way you treat me",
    "your kindness",
    "the way you care",
    "the way it's so easy for you to make me laugh",
    "our sleep calls",
    "just our calls in general",
    "the way he laughs with me",
    "the way he talks to me",
    "the way you help motivate me",
    "your honesty",
    "your loyalty",
    "how strong you are (physically and mentally)",
    "the way we clicked so fast",
    "i found peace in talking to you",
    "i found my best friend in you",
    "i see my future in your eyes",
    "how fun you are",
    "the way you appreciate everything",
    "your gratefulness",
    "you never doubt me",
    "the way you support me",
    "you're always proud of me",
    "you're always by my side",
    "you love me unconditionally",
    "you make me feel enough",
    "you make me feel worth it",
    "you inspire me",
    "you help me become better",
    "you change my view on the world",
    "you stay even when things are hard",
    "you teach me things",
    "you never get disappointed in me",
    "you talk so calmly",
    "you never give up",
    "you make me feel so loved",
    "you make me feel cared",
    "your always there for me",
    "the way you treat people",
    "the way you treat animals",
    "your love for kids",
    "always thinks of me",
    "you try your best",
    "never a bad influence",
    "when we call we don't even have to talk we just like each others presence",
    "never aggressive without a reason",
    "always motivated",
    "has a reason",
    "tries to make people happy",
    "how you see things",
    "make me so happy",
    "give me a reason",
    "your lips",
    "your presence",
    "you keep your promises",
    "your reliable",
    "you help",
    "you're sweet",
    "your kind",
    "your calm even when it's stressful",
    "talking to you feels right",
    "your my soulmate",
    "you give me special memories",
    "i love every moment with you",
    "you communicate",
    "you're committed",
    "you're disciplined",
    "you love me for a lot more than just my body",
    "you believe in me",
    "you believe in yourself",
    "you're mostly positive",
    "you're perfect",
    "never bored of me",
    "never gets annoyed at me",
    "the names you call me",
    "you make me feel special",
    "appreciates small things",
    "your giggles",
    "you always understand",
    "your glow in your eyes",
    "your happy even for the smallest things",
    "you don't care if i have money",
    "never goes to bed angry",
    "calls me when i ask",
    "calls me if i'm upset",
    "helps me when im upset",
    "you don't give up on goals",
    "gives me hope",
    "makes life easier",
    "makes life feel worth it",
    "your my safe place",
    "good mindset",
    "never fail to make me laugh",
    "never fail to make me smile",
    "you compliment me",
    "you compliment my flaws",
    "has good interests",
    "had good intentions",
    "shows me what your capable of",
    "never make me feel insecure",
    "never makes me feel unsafe",
    "you help my mistakes",
    "you let me explain",
    "you feel like home",
    "you're the best gift",
    "makes everything easier",
    "makes life exciting",
    "your little breaths",
    "deals with my attitude",
    "helps with my anxiety",
    "helps my mental health",
    "always supports me",
    "my biggest supporter",
    "you help my weaknesses",
    "always makes an effort",
    "always tries",
    "you trust me",
    "you're trustworthy",
    "keeps my secrets",
    "doesn't hide anything",
    "best singer",
    "never ashamed of me",
    "you're my saviour",
    "makes sure i'm okay",
    "checks up on me",
    "respects me",
    "good behaved",
    "relaxing",
    "you're unique",
    "so thankful",
    "says sorry",
    "has good friends",
    "great music taste",
    "loves the songs i listen to",
    "remembers the small details",
    "you're so cute",
    "so beautiful",
    "so pretty",
    "gorgeous hair",
    "never disrespects",
    "i see a future with you",
    "never rude with me",
    "battles",
    "your love for muay thai",
    "never fails to impress me",
    "impressive",
    "independent",
    "defends me",
    "you always check up",
    "makes sure i've ate",
    "makes sure i sleep",
    "genuine",
    "calms me down",
    "brought back my spark",
    "gives me so much happiness",
    "you actually care",
    "shows you want a future with me",
    "talks about our future together",
    "makes me the happiest person ever",
    "came into my life when i needed it most",
    "i always smile when talking to you",
    "you love me a lot",
    "you're the sweetest boy alive.",
    "your my precious boy",
    "i don't know what i'd do w out you",
    "you always have my back",
    "your knowledge",
    "you're so smart",
    "your teeth",
    "most gorgeous brown eyes",
    "takes me seriously",
    "loves small challenges",
    "talks about us so sweetly",
    "defensive",
    "so full of love",
    "the most strongest love",
    "your tall",
    "never takes advantage of me",
    "you're so generous",
    "you're gentle",
    "always chooses me",
    "i can trust you with my life",
    "you're one of a kind",
    "acts like a little kid sometimes",
    "such a cutie",
    "you've saved me",
    "heart made of pure gold",
    "you stand out",
    "shares everything with me",
    "changes my perspective",
    "you're the reason i'm still going",
    "you're so beautiful",
    "such a cute voice",
    "your always sleepy",
    "usually doesn't take too long to reply",
    "respect's boundaries",
    "i love you for you.",
    "i love your soul",
    "your gentleness",
    "the way you use words",
    "the way you show affection",
    "the way you carry yourself",
    "the way your she's like up",
    "your sweetness",
    "how you mess with your hair",
    "you calm my anxiety",
    "you put my mind a peace",
    "how the quietness isn't awkward",
    "i never have to question you",
    "how \"weird\" things are normal to us",
    "you're brave",
    "how i'm in love",
    "how lucky i am to have you",
    "how grateful i am",
    "how we got so close so fast",
    "no argument can split us up",
    "how we're so comfortable with each other",
    "i can count on you",
    "your my favourite person",
    "make me feel beautiful in and out",
    "makes me feel important",
    "makes the world brighter",
    "our jokes",
    "your goofy side",
    "you're so hardworking",
    "you have the kindest heart",
]


# ============================================================
# ENVOI
# ============================================================
def capitalize_first(text: str) -> str:
    """Met la premiere lettre en majuscule, laisse le reste intact."""
    if not text:
        return text
    return text[0].upper() + text[1:]


def send_message(text: str) -> None:
    """Copie le texte dans le presse-papiers puis colle + envoie."""
    # On utilise le copier/coller pour eviter les problemes
    # de caracteres speciaux et d'accent avec pyautogui.typewrite.
    pyperclip.copy(capitalize_first(text))
    # Ctrl + V (Windows / Linux). Sur Mac il faudrait "command".
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.15)
    pyautogui.press("enter")


def main() -> None:
    total = len(REASONS)
    print(f"=> {total} raisons a envoyer.")
    print(f"=> Reprise a partir de la raison n°{START_INDEX}.")
    print()
    print("Place-toi sur la fenetre WhatsApp Web avec le bon chat ouvert.")
    for i in range(START_DELAY, 0, -1):
        print(f"  Debut dans {i}...", end="\r")
        time.sleep(1)
    print(" " * 30)

    sent = 0
    try:
        for idx, reason in enumerate(REASONS, start=1):
            if idx < START_INDEX:
                continue
            print(f"[{idx:>3}/{total}] {capitalize_first(reason)}")
            send_message(reason)
            sent += 1
            time.sleep(DELAY_BETWEEN_MESSAGES)
    except pyautogui.FailSafeException:
        print("\n[STOP] FailSafe declenche (souris dans un coin).")
        print(f"Derniere raison envoyee : n°{START_INDEX + sent - 1}")
        sys.exit(0)
    except KeyboardInterrupt:
        print("\n[STOP] Interruption clavier.")
        print(f"Derniere raison envoyee : n°{START_INDEX + sent - 1}")
        sys.exit(0)

    print()
    print(f"Termine. {sent} messages envoyes. <3")


if __name__ == "__main__":
    main()
