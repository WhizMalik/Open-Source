import re
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize, sent_tokenize
from heapq import nlargest
from docx import Document

def generate_notes(transcript, num_sentences=5):
    # Tokenize the transcript into sentences
    sentences = sent_tokenize(transcript)

    # Tokenize each sentence into words and remove stopwords
    stop_words = set(stopwords.words('english'))
    word_frequencies = {}
    for sentence in sentences:
        words = word_tokenize(sentence.lower())
        for word in words:
            if word not in stop_words and word.isalnum():
                if word not in word_frequencies:
                    word_frequencies[word] = 1
                else:
                    word_frequencies[word] += 1

    # Calculate sentence scores based on word frequencies
    sentence_scores = {}
    for i, sentence in enumerate(sentences):
        for word in word_tokenize(sentence.lower()):
            if word in word_frequencies:
                if i not in sentence_scores:
                    sentence_scores[i] = word_frequencies[word]
                else:
                    sentence_scores[i] += word_frequencies[word]

    # Get the top N sentences with the highest scores
    summary_sentences = nlargest(num_sentences, sentence_scores, key=sentence_scores.get)

    # Sort the summary sentences in the order they appear in the original transcript
    summary_sentences.sort()

    # Join the summary sentences to form the notes
    notes = " ".join([sentences[i] for i in summary_sentences])

    return notes

# Example usage
transcript = """In the realm of artificial intelligence, machine learning has emerged as a transformative force, reshaping industries and unlocking new possibilities. At its core, machine learning involves the development of algorithms and statistical models that enable computer systems to learn and improve their performance on specific tasks without being explicitly programmed. By leveraging vast amounts of data, these algorithms can identify patterns, make predictions, and adapt to new information, mimicking the way humans learn from experience. The applications of machine learning are far-reaching, spanning across various domains such as healthcare, finance, transportation, and entertainment. In healthcare, machine learning models can assist in diagnosing diseases, predicting patient outcomes, and personalizing treatment plans. By analyzing medical records, imaging data, and genetic information, these models can uncover hidden patterns and provide valuable insights to healthcare professionals. In the financial sector, machine learning algorithms are employed for fraud detection, risk assessment, and algorithmic trading. By analyzing historical data and real-time market trends, these models can identify potential fraudulent activities, assess credit risks, and make data-driven investment decisions. The transportation industry has also witnessed significant advancements through machine learning. Self-driving cars rely on machine learning algorithms to perceive their surroundings, make decisions, and navigate safely on the roads. By continuously learning from sensor data and real-world experiences, these autonomous vehicles can improve their performance and adapt to dynamic environments. Moreover, machine learning has revolutionized the entertainment industry, particularly in the realm of personalized recommendations. Streaming platforms like Netflix and Spotify utilize machine learning algorithms to analyze user preferences, viewing history, and listening habits to provide tailored content recommendations. This enhances user engagement and satisfaction by delivering relevant and enjoyable experiences. However, the rise of machine learning also presents challenges and ethical considerations. As these algorithms become more sophisticated and influential in decision-making processes, it is crucial to address issues such as bias, transparency, and accountability. Ensuring that machine learning models are fair, unbiased, and aligned with human values is of utmost importance to prevent unintended consequences and promote responsible AI development. Furthermore, the need for large, diverse, and representative datasets is paramount to train accurate and reliable machine learning models. Efforts must be made to collect and curate high-quality data while respecting privacy and data protection regulations. Despite the challenges, the future of machine learning holds immense promise. As research advances and computational power increases, we can expect to see even more impressive applications and breakthroughs in the coming years. From personalized medicine and intelligent virtual assistants to autonomous systems and predictive maintenance, machine learning will continue to shape our world in profound ways. As we navigate this exciting frontier, it is essential to foster collaboration between researchers, industry leaders, policymakers, and the public to harness the full potential of machine learning while addressing its challenges and ensuring its responsible deployment for the benefit of society as a whole."""

notes = generate_notes(transcript)
""" print("Generated Notes:")
print(notes) """

#Creating a word doc
document=Document()

document.add_paragraph(notes)

document.save("Generated_notes.docx")