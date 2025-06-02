from google import genai

# Create client using API key
client = genai.Client(api_key="AIzaSyA1hN1uGimOy-IxttnWB9WSvDXx_uvlBok")

# Get user prompt in English
user_prompt = input("Please enter your prompt: ")

# Generate content using the entered prompt
response = client.models.generate_content(
    model="gemini-2.0-flash",
    contents=user_prompt,
)

print("\nGenerated Response:\n")
print(response.text)
