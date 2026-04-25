"""
Module 3: Feature Engineering
TF-IDF + LSA | LDA (sklearn) | SBERT (with TF-IDF fallback)
"""
import numpy as np
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer, CountVectorizer
from sklearn.decomposition import TruncatedSVD, LatentDirichletAllocation
from sklearn.metrics.pairwise import cosine_similarity

# -- TF-IDF ------------------------------------------------------------------
def build_tfidf(df, text_col="processed", max_features=3000):
    print("[Features] Building TF-IDF ...")
    vec = TfidfVectorizer(max_features=max_features, ngram_range=(1,2),
                           sublinear_tf=True, min_df=1)
    mat = vec.fit_transform(df[text_col])
    print(f"[Features] TF-IDF shape: {mat.shape}")
    return mat, vec

def reduce_tfidf(mat, n_components=50):
    nc = min(n_components, mat.shape[1]-1, mat.shape[0]-1)
    svd = TruncatedSVD(n_components=nc, random_state=42)
    reduced = svd.fit_transform(mat)
    print(f"[Features] TF-IDF LSA: {reduced.shape}")
    return reduced, svd

# -- LDA ---------------------------------------------------------------------
def build_lda(df, tokens_col="tokens", n_topics=6):
    print(f"[Features] Building LDA ({n_topics} topics) ...")
    docs = [" ".join(t) for t in df[tokens_col]]
    cv = CountVectorizer(max_features=2000, min_df=1)
    X  = cv.fit_transform(docs)
    lda = LatentDirichletAllocation(n_components=n_topics, random_state=42,
                                     max_iter=20, learning_method="batch")
    mat = lda.fit_transform(X)
    # store vocab for top-word retrieval
    lda._vocab = cv.get_feature_names_out()
    print(f"[Features] LDA matrix: {mat.shape}")
    return mat, lda, cv, X

def get_lda_top_words(lda, n_words=10):
    vocab = lda._vocab
    topics = []
    for i, comp in enumerate(lda.components_):
        top = np.argsort(comp)[::-1][:n_words]
        topics.append("Topic %d: " % i + ", ".join(vocab[top]))
    return topics

# -- Sentence-BERT (with fallback) -------------------------------------------
def build_sbert(df, text_col="text", model_name="all-MiniLM-L6-v2"):
    try:
        from sentence_transformers import SentenceTransformer
        print(f"[Features] Building SBERT ({model_name}) ...")
        model = SentenceTransformer(model_name)
        emb = model.encode(df[text_col].tolist(), show_progress_bar=True,
                           convert_to_numpy=True)
        print(f"[Features] SBERT: {emb.shape}")
        return emb
    except Exception as e:
        print(f"[Features] SBERT unavailable ({e}). Using TF-IDF LSA as proxy.")
        mat, _ = build_tfidf(df, text_col="processed", max_features=3000)
        emb, _ = reduce_tfidf(mat, n_components=64)
        return emb
    
def compute_topic_coherence(lda_model, vectorizer, texts, top_n=10):
    """
    Simple coherence using word co-occurrence in corpus (PMI-style proxy)
    """
    import numpy as np

    # Get vocabulary and document-term matrix
    vocab = vectorizer.get_feature_names_out()
    X = vectorizer.transform(texts) 

    topic_scores = []

    for topic in lda_model.components_:
        top_idx = topic.argsort()[::-1][:top_n]

        # get columns for top words
        word_vectors = X[:, top_idx].toarray()

        pair_scores = []

        for i in range(len(top_idx)):
            for j in range(i + 1, len(top_idx)):
                wi = word_vectors[:, i]
                wj = word_vectors[:, j]

                # co-occurrence: docs where both appear
                co_occur = np.sum((wi > 0) & (wj > 0))
                freq_i = np.sum(wi > 0)
                freq_j = np.sum(wj > 0)

                if co_occur > 0:
                    score = co_occur / (freq_i + freq_j)
                    pair_scores.append(score)

        if pair_scores:
            topic_scores.append(np.mean(pair_scores))

    return float(np.mean(topic_scores)) if topic_scores else 0.0

# -- Master builder -----------------------------------------------------------
def build_all_features(df, n_topics=6, tfidf_components=50):
    tfidf_raw, tfidf_vec = build_tfidf(df)
    tfidf_lsa, _         = reduce_tfidf(tfidf_raw, n_components=tfidf_components)
    lda_mat, lda_model, lda_cv, lda_corpus = build_lda(df, n_topics=n_topics)
    sbert_mat = build_sbert(df)
    coherence = compute_topic_coherence(
    lda_model,
    lda_cv,
    df["processed"]
)
    print(f"[LDA] Topic Coherence: {coherence:.4f}")
    return {
        "tfidf_raw":  tfidf_raw,
        "tfidf_lsa":  tfidf_lsa,
        "lda":        lda_mat,
        "sbert":      sbert_mat,
        "lda_model":  lda_model,
        "lda_cv":     lda_cv,
        "lda_corpus": lda_corpus,
        "tfidf_vec":  tfidf_vec,
        "lda_coherence": coherence
    }

def find_best_topics(df, topic_range=range(2, 10)):
    scores = {}
    for k in topic_range:
        lda_mat, lda_model, lda_cv, _ = build_lda(df, n_topics=k)
        score = compute_topic_coherence(lda_model, lda_cv, df["processed"])
        scores[k] = score
        print(f"[LDA] topics={k} coherence={score:.4f}")
    best_k = max(scores, key=scores.get)
    print(f"[LDA] Best topics based on coherence: {best_k}")
    return best_k, scores
