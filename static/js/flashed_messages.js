// flash_messages.js

// Function to style flashed messages
function styleFlashMessages() {
    // Get all flashed messages with the 'alert-danger' class
    const dangerMessages = document.querySelectorAll('.alert-danger');

    // Iterate through danger messages and apply styles
    dangerMessages.forEach(message => {
        // Style the message as needed
        message.style.color = 'red';
    });
}

// Call the function to style flashed messages when the page loads
document.addEventListener('DOMContentLoaded', styleFlashMessages);
