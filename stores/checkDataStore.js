import axios from "axios"
import { defineStore } from "pinia"
import { ref } from "vue"

export const checkDataStore = defineStore("checkData", () => {
  const data = ref(null)
  const loadingData = ref(false)

  const checkFile = async (file) => {
    loadingData.value = true
    const formData = new FormData()
    formData.append("file", file)
    formData.append("name", file.name)
    try {
      const response = await axios.post(
        `https://imaginative-conkies-4a8f20.netlify.app/api/checkfile`,
        formData,
        {
          headers: {
            "Content-Type": "multipart/form-data",
          },
        }
      )
      data.value = [...response.data]
    } catch (error) {
      console.log(error)
    } finally {
      loadingData.value = false
    }
  }

  return {
    data,
    loadingData,
    checkFile,
  }
})
