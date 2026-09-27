"""Duplicate-Detection und Tagging-Logik."""
import asyncio
from typing import Dict, List


async def get_all_documents(paperless) -> List[Dict]:
    """Alle Dokumente von Paperless abrufen."""
    print("📚 Lade Dokumente...")

    documents = []
    page = 1

    while True:
        try:
            # Nutze die pypaperless API
            docs = await paperless.getDocuments(page=page, page_size=100)

            if not docs:
                break

            documents.extend(docs)
            print(f"   ... {len(documents)} Dokumente geladen", end="\r")
            page += 1

            # Kleine Pause, um Paperless nicht zu überlasten
            await asyncio.sleep(0.1)

        except Exception as e:
            print(f"\n❌ Fehler beim Abrufen der Dokumente: {e}")
            raise

    print(f"\n✓ Insgesamt {len(documents)} Dokumente geladen\n")
    return documents


async def ensure_tag_exists(paperless, tag_name: str, tag_color: str) -> int:
    """Stellt sicher, dass das Tag existiert. Erstellt es ggf."""
    print(f"🏷️  Prüfe Tag '{tag_name}'...")

    try:
        # Versuche Tag zu laden
        tags = await paperless.getTags()

        for tag in tags:
            if tag.name == tag_name:
                print(f"   ✓ Tag '{tag_name}' existiert (ID: {tag.id})\n")
                return tag.id

        # Tag existiert nicht → erstellen
        print(f"   • Tag '{tag_name}' existiert nicht, erstelle...")
        new_tag = await paperless.createTag(name=tag_name, color=tag_color)
        print(f"   ✓ Tag erstellt (ID: {new_tag.id})\n")
        return new_tag.id

    except Exception as e:
        print(f"❌ Fehler beim Tag-Management: {e}")
        raise


async def mark_duplicates_async(
    paperless,
    tag_name: str = "hat_duplikate",
    tag_color: str = "#FF6B6B"
) -> None:
    """
    Markiert alle Dokumente mit Duplikaten mit dem angegebenen Tag.

    Args:
        paperless: Initialized Paperless instance
        tag_name: Name des Tags für Duplikate
        tag_color: Farbe des Tags (Hex, z.B. #FF6B6B für Rot)
    """

    # Tag vorbereiten
    tag_id = await ensure_tag_exists(paperless, tag_name, tag_color)

    # Dokumente laden
    documents = await get_all_documents(paperless)

    # Duplikate finden
    docs_with_duplicates = [
        (doc, doc.duplicate_count)
        for doc in documents
        if hasattr(doc, 'duplicate_count') and doc.duplicate_count > 0
    ]

    if not docs_with_duplicates:
        print("✓ Keine Duplikate gefunden!")
        return

    print(f"🔎 Gefunden: {len(docs_with_duplicates)} Dokumente mit Duplikaten\n")

    tagged_count = 0
    already_tagged_count = 0
    error_count = 0

    for doc, dup_count in docs_with_duplicates:
        doc_id = doc.id
        title = doc.title[:50] if hasattr(doc, 'title') else f"ID {doc_id}"

        try:
            # Prüfe, ob Tag bereits vorhanden
            tag_ids = doc.tags if hasattr(doc, 'tags') else []

            if tag_id in tag_ids:
                print(f"  → ID {doc_id}: {title}... ({dup_count}× dup, bereits getaggt)")
                already_tagged_count += 1
            else:
                # Tag hinzufügen
                tag_ids.append(tag_id)
                await paperless.updateDocument(doc_id, tags=tag_ids)
                print(f"  ✓ ID {doc_id}: {title}... ({dup_count}× dup)")
                tagged_count += 1

            # Kleine Pause
            await asyncio.sleep(0.05)

        except Exception as e:
            print(f"  ✗ ID {doc_id}: Fehler ({str(e)[:50]})")
            error_count += 1

    # Summary
    print(f"\n📊 Zusammenfassung:")
    print(f"   ✓ {tagged_count} Dokumente neu getaggt")
    print(f"   → {already_tagged_count} waren bereits getaggt")
    if error_count:
        print(f"   ✗ {error_count} Fehler")
    print(f"\n💡 Filtere in der Paperless-UI: tag:{tag_name}")
