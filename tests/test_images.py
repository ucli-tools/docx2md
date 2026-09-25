"""
Tests for the image processing module.
"""

from docx2md.processors.images import extract_and_process_images


def _extract(tmp_path, markdown):
    img = tmp_path / "img"
    media = img / "media"
    media.mkdir(parents=True)
    for name in ("image1.png", "image2.png"):
        (media / name).write_bytes(b"png")
    md = tmp_path / "book.md"
    md.write_text(markdown, encoding="utf-8")
    stats = extract_and_process_images(md, md, images_dir=str(img), optimize=False)
    return img, md.read_text(encoding="utf-8"), stats


def test_images_moved_up_and_media_removed(tmp_path):
    """pandoc's img/media/ copies are removed once each image is in img/."""
    img, text, stats = _extract(
        tmp_path, "![](img/media/image1.png)\n\n![](img/media/image2.png)\n"
    )
    assert sorted(p.name for p in img.iterdir()) == ["image1.png", "image2.png"]
    assert not (img / "media").exists()
    assert "img/media" not in text
    assert stats["images_processed"] == 2


def test_image_referenced_twice(tmp_path):
    """A second reference to the same image still resolves."""
    img, text, stats = _extract(
        tmp_path, "![](img/media/image1.png)\n\n![](img/media/image1.png)\n"
    )
    assert (img / "image1.png").exists()
    assert stats["images_failed"] == 0
    assert text.count("img/image1.png") == 2


def test_unreferenced_extraction_kept(tmp_path):
    """Only images that were copied up are removed from media/."""
    img, _, _ = _extract(tmp_path, "![](img/media/image1.png)\n")
    assert (img / "media" / "image2.png").exists()
    assert not (img / "media" / "image1.png").exists()
