function handleFormSubmit(event) {
    event.preventDefault();
    const file = fileInput.files[0];
    if (!file) return;
      

    showLoader();
  
    const formData = new FormData();
    formData.append('xlsxFile', file);
  
    axios.post('/generate', formData,{ responseType: 'arraybuffer' })
      .then(response => {
        hideLoader();
        if (!response || !response.data){
          showResult('Error: No Reponse Data');
          return;
        }
        showResult('Success: ');
        const fileData = new Blob([response.data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const fileUrl = URL.createObjectURL(fileData);
        
        // Create a link element and trigger a download
        const downloadLink = document.createElement('a');
        downloadLink.href = fileUrl;
        downloadLink.download = 'converted_file.xlsx';
        downloadLink.click();
      })
      .catch(error => {
        hideLoader();
        showResult('Error: ' + error.response.data);
      });
  }
  
  function showLoader() {
    loader.classList.remove('d-none');
  }
  
  function hideLoader() {
    loader.classList.add('d-none');
  }
  
  function showResult(message) {
    resultContainer.innerText = message;
  }
  