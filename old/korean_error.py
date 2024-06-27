from hanspell import spell_checker

def check_spelling(text):
    result = spell_checker.check(text)
    errors = result.as_dict()
    
    corrected_text = errors['checked']
    original_text = errors['original']
    
    # Find differences between original and corrected text
    if original_text != corrected_text:
        corrections = []
        original_words = original_text.split()
        corrected_words = corrected_text.split()
        
        for orig, corr in zip(original_words, corrected_words):
            if orig != corr:
                corrections.append((orig, corr))
        
        return corrections
    else:
        return []

text_to_check = "안녕하11세요"

errors = check_spelling(text_to_check)

if errors:
    for orig, corr in errors:
        print(f"오타: {orig} -> 수정 제안: {corr}")
else:  

    print("오타가 없습니다.")
