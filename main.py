from src.run import run


if __name__ == '__main__':
    zotero_api_key = "A3sxgk3eNYO7xmAIldd0kgkB"
    zotero_id = "7367571"
    word_file_path = r"D:\Zizi_Master_student\thesis\zotero\link-zotero-citation-bibliography\test.docx"
    new_file_path = r"D:\Zizi_Master_student\thesis\zotero\link-zotero-citation-bibliography\result\test_result.docx"
    run(word_file_path, new_file_path, zotero_id, zotero_api_key, isNumbered=True, setColor=16711680, noUnderLine=True)
