const userInput = document.getElementById('user-input');
const sendButton = document.getElementById('send-button');
const chatContainer = document.getElementById('chat-container');

sendButton.addEventListener('click', sendMessage);
userInput.addEventListener('keyup', function(event) {
  if (event.key === 'Enter') {
    sendMessage();
  }
});

function sendMessage() {
  const message = userInput.value.trim();
  if (message === "") return;


  const userMessageDiv = document.createElement('div');
  userMessageDiv.classList.add('message', 'user');
  userMessageDiv.textContent = message;
  chatContainer.appendChild(userMessageDiv);

  userInput.value = '';  

  
  fetch('/chat', {  
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ message: message })
  })
  .then(response => response.json())
  .then(data => { 
    const botMessageDiv = document.createElement('div');
    botMessageDiv.classList.add('message', 'bot');
    botMessageDiv.textContent = data.message; 
    chatContainer.appendChild(botMessageDiv);
    chatContainer.scrollTop = chatContainer.scrollHeight;
  })
  .catch(error => console.error('Error:', error));


  chatContainer.scrollTop = chatContainer.scrollHeight;
}