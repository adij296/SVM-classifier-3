# Question Label Classifier

This project processes raw question labels and their descriptions from a Word document, predicting and turning them into standardized canonical labels using a trained SVM model, and compares a separate list of ground truth labels against the raw labels to determine whether each is present. There are two variations: one that simply classifies all the labels, and another than will mark a lubel as unknown if the model is not fully sure, rather than outputting an incorrect label.

The output is two Word documents:
- ProcessedLabels.docx – Contains the predicted canonical labels along with the unchanged label descriptions.

- GroundTruth.docx – Contains ground truth labels with a "Found/Not Found" status depending on whether they appear in the raw labels.

## Features

- Converts raw, unstructured labels into a standardized canonical format.
- Describes labels as unknown if they are not recognized
- Uses a character-level TF-IDF vectorizer and SVM classifier for flexible recognition of label variations.
- Maintains a separate ground truth list that automatically updates with "Found/Not Found" results.
- Works with Microsoft Word .docx files for both input and output.

## Requirements

- Python 3.8+
- Install dependencies with:
  - pip install scikit-learn pandas python-docx openpyxl

## File Structure

- sample - Input.docx: Input Word document containing raw labels and a description a line below.

- GroundTruth.docx: Input/output Word document containing ground truth labels.

  - On first run, this may contain a simple header ("Ground Truth") followed by labels.

  - The script will convert this into a table and maintain it going forward.

- clean_full_training_dataset.xlsx: Training dataset containing raw_label, description and canonical_label columns.

- ProcessedLabels.docx: Output Word document with raw labels and predicted canonical labels, unchaged descriptions (and labels marked as UNKNOWN, if applicable).
  
- LabelClassifier.py: Code for classifier that classifies all raw labels
  
- LabelClassifierWithUnknowns.py: Code for classifier that classifies labels and marks unrecognized labels as unknown

## Usage

Place your raw labels into sample - Input.docx, one label followed by description on individual lines (a prefilled sample sample - Input.docx is available in the repository)

Prepare GroundTruth.docx with a table, the first column (with a header of "Ground Truth") having all the base labels, and a blank second column (with a header of "Found/Not Found"). There is also a sample GroundTruth.docx within the repository.

Run either scripts, with or without unknowns. The script will:

- Train the SVM model on your Excel dataset.

- Generate ProcessedLabels.docx with predictions.

- Update GroundTruth.docx into a table, filling the "Found/Not Found" column.

## Adding New Ground Truths

- The GroundTruth.docx will be a table.

- To add more labels, insert new rows into the Ground Truth column of the table.

- Rerun the script to update statuses.
