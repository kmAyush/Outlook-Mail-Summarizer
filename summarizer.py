from transformers import BartTokenizer, BartForConditionalGeneration, BartConfig

checkpoint = "facebook/bart-large-cnn"
tokenizer = BartTokenizer.from_pretrained(checkpoint)
model = BartForConditionalGeneration.from_pretrained(checkpoint)

def summarize(f):
      
  sentences = f

  # initialize
  length = 0
  chunk = ""
  chunks = []
  count = -1
  for sentence in sentences:
    count += 1
    combined_length = len(tokenizer.tokenize(sentence)) + length # add the no. of sentence tokens to the length counter

    if combined_length  <= tokenizer.max_len_single_sentence: # if it doesn't exceed
      chunk += sentence + " " # add the sentence to the chunk
      length = combined_length # update the length counter

      # if it is the last sentence
      if count == len(sentences) - 1:
        chunks.append(chunk.strip()) # save the chunk
      
    else: 
      chunks.append(chunk.strip()) # save the chunk
      
      # reset 
      length = 0 
      chunk = ""

      # take care of the overflow sentence
      chunk += sentence + " "
      length = len(tokenizer.tokenize(sentence))
  
  # inputs to the model
  inputs = [tokenizer(chunk, return_tensors="pt") for chunk in chunks]

  for input in inputs:
    output = model.generate(**input)
    print(tokenizer.decode(*output, skip_special_tokens=True))

def summarizer(text):
  inputs = tokenizer([text], return_tensors = 'pt')
  summary_ids = model.generate(inputs['input_ids'], max_length=1000, early_stopping=False)
  output = [tokenizer.decode(g, skip_special_tokens=True) for g in summary_ids]
  return output[0]
