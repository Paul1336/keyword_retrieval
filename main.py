import os
import nltk
from nltk.corpus import stopwords
from nltk.stem.snowball import SnowballStemmer
import spacy
from collections import Counter
from nltk.stem.snowball import FrenchStemmer
import jieba
import pandas as pd

nltk.download('stopwords')


def analyze_russian_texts(folder_path, top_n=20):
    nlp = spacy.load('ru_core_news_sm')
    russian_stopwords = set(stopwords.words('russian'))
    stemmer = SnowballStemmer("russian")

    all_tokens = []

    for file_name in os.listdir(folder_path):
        if file_name.endswith('.txt'):
            file_path = os.path.join(folder_path, file_name)
            with open(file_path, 'r', encoding='utf-8') as file:
                text = file.read()

            doc = nlp(text)
            tokens = [token.text.lower() for token in doc if token.is_alpha]
            tokens_no_stopwords = [token for token in tokens if token not in russian_stopwords]
            stemmed_tokens = [stemmer.stem(token) for token in tokens_no_stopwords]
            all_tokens.extend(stemmed_tokens)

    word_freq = Counter(all_tokens)
    return word_freq.most_common(top_n)


def analyze_french_texts(folder_path, top_n=20):
    nlp = spacy.load('fr_core_news_sm')
    french_stopwords = set(stopwords.words('french'))
    stemmer = FrenchStemmer()

    all_tokens = []

    for file_name in os.listdir(folder_path):
        if file_name.endswith('.txt'):
            file_path = os.path.join(folder_path, file_name)
            with open(file_path, 'r', encoding='utf-8') as file:
                text = file.read()

            doc = nlp(text)
            tokens = [token.text.lower() for token in doc if token.is_alpha]
            tokens_no_stopwords = [token for token in tokens if token not in french_stopwords]
            stemmed_tokens = [stemmer.stem(token) for token in tokens_no_stopwords]
            all_tokens.extend(stemmed_tokens)

    word_freq = Counter(all_tokens)
    return word_freq.most_common(top_n)



def analyze_english_texts(folder_path, top_n=20):
    nlp = spacy.load('en_core_web_sm')
    english_stopwords = set(stopwords.words('english'))
    stemmer = SnowballStemmer("english")

    all_tokens = []

    for file_name in os.listdir(folder_path):
        if file_name.endswith('.txt'):
            file_path = os.path.join(folder_path, file_name)
            with open(file_path, 'r', encoding='utf-8') as file:
                text = file.read()

            doc = nlp(text)
            tokens = [token.text.lower() for token in doc if token.is_alpha]
            tokens_no_stopwords = [token for token in tokens if token not in english_stopwords]
            stemmed_tokens = [stemmer.stem(token) for token in tokens_no_stopwords]
            all_tokens.extend(stemmed_tokens)

    word_freq = Counter(all_tokens)
    return word_freq.most_common(top_n)

def analyze_chinese_texts(folder_path, top_n=20):
    default_stopwords = set(["的", "了", "和", "是", "我", "有", "在", "不", "就", "人", "都", "一", "一个", "上", "也", "很", "到", "说", "而", "以", "他", "她"])
    all_tokens = []
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.txt'): 
            file_path = os.path.join(folder_path, file_name)
            with open(file_path, 'r', encoding='utf-8') as file:
                text = file.read()

            tokens = jieba.cut(text)
            # print("token", tokens)
            tokens_no_stopwords = [token for token in tokens if token not in default_stopwords and len(token.strip()) > 1]
            all_tokens.extend(tokens_no_stopwords)

    word_freq = Counter(all_tokens)
    return word_freq.most_common(top_n)

def save_results_to_file(file_path, result):
    with open(file_path, 'w', encoding='utf-8') as f:
        for rank, (word, freq) in enumerate(result, start=1):
            f.write(f"{rank} - {word}: {freq}\n")


def save_results_to_excel(file_path, results_dict):
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        for language, result in results_dict.items():
            df = pd.DataFrame(result, columns=["Word", "Frequency"])
            df.insert(0, "Rank", range(1, len(df) + 1))
            df.to_excel(writer, sheet_name=language, index=False)

if __name__ == "__main__":
    results = {}

    folder_path = './TW'
    chinese_result = analyze_chinese_texts(folder_path, top_n=200)
    results['Chinese'] = chinese_result
    save_results_to_file('./chinese_results.txt', chinese_result)
    print("saved to chinese_results.txt")

    folder_path = './RF'
    russian_result = analyze_russian_texts(folder_path, top_n=200)
    results['Russian'] = russian_result
    save_results_to_file('./russian_results.txt', russian_result)
    print("saved to russian_results.txt")

    folder_path = './FR'
    french_result = analyze_french_texts(folder_path, top_n=200)
    results['French'] = french_result
    save_results_to_file('./french_results.txt', french_result)
    print("saved to french_results.txt")

    folder_path = './US'
    english_result = analyze_english_texts(folder_path, top_n=200)
    results['English'] = english_result
    save_results_to_file('./english_results.txt', english_result)
    print("saved to english_results.txt")

    save_results_to_excel('./word_frequency_results.xlsx', results)
    print("saved to word_frequency_results.xlsx")