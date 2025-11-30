from datetime import datetime
from agents import get_chat_history
from fastapi import Query
import uvicorn
from fastapi import FastAPI, Form, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from agents import ask_something, PPTAgentResp, PPT_PPT_FILE
import os
from dotenv import load_dotenv
load_dotenv()


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


@app.get("/")
async def root():
    return {"message": "Welcome to the PPT Generation API",
            "timestamp": datetime.utcnow().isoformat() + "Z"}


@app.post("/generate")
async def generate_presentation(topic: str = Form(...)):
    """
    Receives a topic, generates a presentation, and returns it.
    """
    session_id = 1  # Using a single session for now
    try:
        # This function will trigger the agent chain which creates and saves the pptx file.
        result = ask_something(session_id, topic)

        if result.ppt_generated:
            if os.path.exists(PPT_PPT_FILE):
                # Return the pptx file as a response
                return FileResponse(
                    PPT_PPT_FILE,
                    media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    filename="output.pptx",
                    headers={"status": "true", "content": str(result.content)}
                )
            else:
                raise HTTPException(
                    status_code=500, detail="Presentation file was not found after generation.")
        else:
            return {"status": False, "content": result.content}

    except Exception as e:
        raise HTTPException(
            status_code=500, detail=f"An error occurred: {str(e)}")


# New endpoint to get session history by session_id


@app.get("/session_history")
async def get_session_history_endpoint(session_id: int = Query(..., description="Session ID")):
    """
    Returns the session history for a given session_id.
    """
    try:
        history = get_chat_history(session_id)
        return {"session_id": session_id, "history": history}
    except Exception as e:
        raise HTTPException(
            status_code=500, detail=f"Could not fetch session history: {str(e)}")


if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
