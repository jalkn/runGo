import flask as fk
import os
import app

AICHAT = os.environ.get("GOOGLE_API_KEY")
if AICHAT is None:
    raise ValueError("aichat not set.")

@app.route("/aichat", methods=["POST"])
def aichat():
    user_message = fk.request.form.get("user_message")

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {AICHAT}" 
    }

    request_body = {
        "queryInput": {
            "text": {
                "text": user_message
            }
        }
    }

    try:
        response = fk.requests.post(AICHAT, headers=headers, json=request_body)
        response.raise_for_status()  
        response_json = response.json()
        bot_message = response_json["queryResult"]["fulfillmentText"] 

    except fk.requests.exceptions.RequestException as e:
        print(f"Error with Gemini API: {e}")
        bot_message = "I'm having trouble right now. Please try again later."


    messages = fk.session.get('messages', [])
    messages.append({"role": "user", "content": user_message})
    messages.append({"role": "assistant", "content": bot_message})
    fk.session['messages'] = messages

    return fk.jsonify({"messages": messages[-2:]})