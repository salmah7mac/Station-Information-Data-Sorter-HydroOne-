import pandas as pd
from sklearn.pipeline import Pipeline
from sklearn.model_selection import train_test_split, cross_val_predict, cross_val_score
from sklearn.preprocessing import LabelEncoder
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics import classification_report, accuracy_score
from sklearn.ensemble import RandomForestClassifier

#from google.colab import files
#uploaded = files.upload()

# Load the CSV file
file_path = 'input.csv'
df = pd.read_csv(file_path)

# Encode the target variable
label_encoder = LabelEncoder()
df['Category_encoded'] = label_encoder.fit_transform(df['Category'])

# Split the data into features and target variable
X = df['FileName']
y = df['Category_encoded']

# Manually split the data into training and testing sets
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# Create a pipeline
pipeline = Pipeline([
    ('tfidf', TfidfVectorizer()),
    ('classifier', RandomForestClassifier(random_state=42))
])

# Perform cross-validation on the training set and predict on the testing set
pipeline.fit(X_train, y_train)
y_pred = pipeline.predict(X_test)

# Get the unique labels from the entire dataset
unique_labels = label_encoder.classes_

# Evaluate the model
accuracy = accuracy_score(y_test, y_pred)
report = classification_report(y_test, y_pred, labels=range(len(unique_labels)), target_names=unique_labels)

print("Accuracy:", accuracy)
print("Classification Report:\n", report)



# Train the pipeline on the entire dataset
pipeline.fit(X, y)

# Load the new dataset for prediction
new_file_path = 'file.csv'
df_new = pd.read_csv(new_file_path)

# Predict on the new dataset
X_new = df_new['FileName']
predictions = pipeline.predict(X_new)

# Decode the predictions to original category names
df_new['Predictions'] = label_encoder.inverse_transform(predictions)

# Export the result to a new CSV file
output_file_path = 'data.csv'
df_new.to_csv(output_file_path, index=False)

# Display the result
df_new.head()