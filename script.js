
// Play audio and video when clicked
document.getElementById('audio1').addEventListener('play', () => {
    console.log('Playing Audio 1');
});
document.getElementById('audio2').addEventListener('play', () => {
    console.log('Playing Audio 2');
});
document.getElementById('video1').addEventListener('play', () => {
    console.log('Playing Video');
});

// Handle form submission
document.getElementById('contactForm').addEventListener('submit', function (e) {
    e.preventDefault();

    // Get form data
    const name = document.getElementById('name').value;
    const email = document.getElementById('email').value;
    const message = document.getElementById('message').value;

    // Create Excel data
    const data = [
        ['Name', 'Email', 'Message'],
        [name, email, message]
    ];

    // Create worksheet and workbook
    const worksheet = XLSX.utils.aoa_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Responses');

    // Export workbook
    XLSX.writeFile(workbook, 'ContactResponses.xlsx');

    // Reset form and notify user
    document.getElementById('contactForm').reset();
    alert('Your response has been saved as an Excel file.');
});