import uvicorn
from fastapi import FastAPI, Form, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from agents import ask_something
import os

app = FastAPI()

# Configure CORS
origins = [
    "http://localhost:3000",  # Allow frontend origin
    "https://ocean-ai-ppt-generator-5883.vercel.app",
    "https://ocean-ai-ppt-generator-tbhc.vercel.app",
    "https://ocean-ai-ppt-generator.vercel.app/",
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

PPT_FILE_PATH = "/tmp/output.pptx"

@app.post("/generate")
async def generate_presentation(topic: str = Form(...)):
    """
    Receives a topic, generates a presentation, and returns it.
    """
    session_id = 1 # Using a single session for now
    try:
        # This function will trigger the agent chain which creates and saves the pptx file.
        result = ask_something(session_id, f"Generate a ppt on topic: {topic}")

        if "done" in result.lower() and os.path.exists(PPT_FILE_PATH):
             # The agent is expected to save the file as `output.pptx` in the project root.
            return FileResponse(
                path=PPT_FILE_PATH,
                filename="presentation.pptx",
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        else:
            # If the agent did not confirm completion or the file doesn't exist
            raise HTTPException(status_code=500, detail=f"Backend agent failed to generate the presentation. Agent response: {result}")

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"An error occurred: {str(e)}")

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
