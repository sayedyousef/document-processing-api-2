<template>
  <div class="max-w-4xl mx-auto mt-6 p-6 bg-white rounded-lg shadow-md">
    <h3 class="text-lg font-semibold mb-4">Results Ready</h3>

    <!-- Add refresh button -->
    <button
      @click="refreshResults"
      class="mb-3 px-3 py-1 bg-gray-100 text-gray-600 text-sm rounded hover:bg-gray-200"
    >
      Refresh Results
    </button>

    <div class="space-y-3">
      <div
        v-for="(result, index) in currentResults"
        :key="index"
        class="flex justify-between items-center p-3 bg-gray-50 rounded-lg"
      >
        <div class="flex items-center space-x-3">
          <svg class="h-5 w-5 text-green-500" fill="currentColor" viewBox="0 0 20 20">
            <path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clip-rule="evenodd" />
          </svg>
          <div>
            <p class="font-medium">{{ result.filename }}</p>
            <p class="text-sm text-gray-600">
              {{ formatFileSize(result.size) }}
              <span v-if="result.type === 'application/zip'" class="ml-2 text-blue-600">
                (Multiple files packaged)
              </span>
            </p>
          </div>
        </div>

        <div class="flex items-center space-x-2">
          <span v-if="result.error" class="text-xs text-red-500">{{ result.error }}</span>
          <template v-else>
            <!-- Preview HTML button (for HTML files inside ZIP) -->
            <button
              v-if="result.type === 'application/zip'"
              @click="previewHtml()"
              class="px-3 py-1 bg-purple-500 text-white text-sm rounded hover:bg-purple-600 transition-colors"
            >
              Preview HTML
            </button>
            <button
              @click="downloadFileWithRefresh(index)"
              class="px-3 py-1 bg-blue-500 text-white text-sm rounded hover:bg-blue-600 transition-colors"
            >
              {{ result.type === 'application/zip' ? 'Download ZIP' : 'Download' }}
            </button>
          </template>
        </div>
      </div>
    </div>

    <!-- HTML Preview Section -->
    <div v-if="showPreview" class="mt-6 border-t pt-6">
      <div class="flex justify-between items-center mb-4">
        <h4 class="text-lg font-semibold">HTML Preview with Rendered Equations</h4>
        <button
          @click="showPreview = false"
          class="text-gray-500 hover:text-gray-700"
        >
          Close
        </button>
      </div>

      <!-- Copy HTML Button -->
      <div class="mb-4 flex space-x-2">
        <button
          @click="copyHtmlToClipboard"
          class="px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600 transition-colors"
        >
          {{ copySuccess ? 'Copied!' : 'Copy HTML to Clipboard' }}
        </button>
        <button
          @click="showHtmlSource = !showHtmlSource"
          class="px-4 py-2 bg-gray-500 text-white rounded hover:bg-gray-600 transition-colors"
        >
          {{ showHtmlSource ? 'Hide Source' : 'Show Source' }}
        </button>
      </div>

      <!-- HTML Source Textarea -->
      <div v-if="showHtmlSource" class="mb-4">
        <label class="block text-sm font-medium text-gray-700 mb-2">HTML Source (copy this):</label>
        <textarea
          ref="htmlSourceTextarea"
          v-model="htmlContent"
          class="w-full h-48 p-3 border border-gray-300 rounded-md font-mono text-sm bg-gray-50"
          dir="ltr"
          readonly
        ></textarea>
      </div>

      <!-- Rendered Preview -->
      <div class="border rounded-lg overflow-hidden">
        <div class="bg-gray-100 px-4 py-2 text-sm text-gray-600">
          Rendered Preview (with MathJax equations)
        </div>
        <iframe
          ref="previewFrame"
          class="w-full bg-white"
          style="min-height: 500px; height: 70vh;"
          sandbox="allow-scripts allow-same-origin"
        ></iframe>
      </div>
    </div>

    <button
      @click="downloadAll"
      v-if="currentResults.length > 1 && !hasErrors"
      class="mt-4 w-full px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600 transition-colors font-semibold"
    >
      Download All as ZIP
    </button>

    <button
      @click="$emit('reset')"
      class="mt-3 w-full px-4 py-2 bg-gray-200 text-gray-700 rounded hover:bg-gray-300 transition-colors"
    >
      Process More Documents
    </button>
  </div>
</template>

<script>
import { ref, computed, onMounted, nextTick } from 'vue'
import axios from 'axios'
import JSZip from 'jszip'
import { API_BASE_URL } from '../config'

export default {
  props: ['results', 'jobId'],
  emits: ['reset'],
  setup(props) {
    // Use local state for results that can be refreshed
    const currentResults = ref(props.results)

    // HTML Preview state
    const showPreview = ref(false)
    const showHtmlSource = ref(false)
    const htmlContent = ref('')
    const copySuccess = ref(false)
    const previewFrame = ref(null)
    const htmlSourceTextarea = ref(null)

    const hasErrors = computed(() => {
      return currentResults.value.some(r => r.error)
    })

    // FORMAT FILE SIZE FUNCTION
    const formatFileSize = (bytes) => {
      console.log('[ResultDownload] Formatting file size:', bytes)
      if (!bytes || bytes === 0) return '0 Bytes'
      const k = 1024
      const sizes = ['Bytes', 'KB', 'MB', 'GB']
      const i = Math.floor(Math.log(bytes) / Math.log(k))
      return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i]
    }

    // Refresh results from backend
    const refreshResults = async () => {
      console.log('[ResultDownload] Refreshing results for job:', props.jobId)
      try {
        const response = await axios.get(`${API_BASE_URL}/api/status/${props.jobId}`)
        if (response.data.results) {
          currentResults.value = response.data.results
          console.log('[ResultDownload] Results refreshed:', currentResults.value)
        }
      } catch (error) {
        console.error('[ResultDownload] Failed to refresh results:', error)
      }
    }

    // Preview HTML from ZIP
    const previewHtml = async () => {
      try {
        console.log('[ResultDownload] Loading HTML preview...')

        // Download the ZIP file
        const response = await axios.get(
          `${API_BASE_URL}/api/download/${props.jobId}`,
          { responseType: 'arraybuffer' }
        )

        // Extract HTML from ZIP
        const zip = await JSZip.loadAsync(response.data)

        // Find the HTML file in the ZIP
        let htmlFile = null
        for (const filename of Object.keys(zip.files)) {
          if (filename.endsWith('.html')) {
            htmlFile = zip.files[filename]
            break
          }
        }

        if (!htmlFile) {
          alert('No HTML file found in the ZIP')
          return
        }

        // Get HTML content
        const html = await htmlFile.async('string')
        htmlContent.value = html

        // Show preview
        showPreview.value = true

        // Wait for DOM update then set iframe content
        await nextTick()
        if (previewFrame.value) {
          const iframeDoc = previewFrame.value.contentDocument || previewFrame.value.contentWindow.document
          iframeDoc.open()
          iframeDoc.write(html)
          iframeDoc.close()
        }

        console.log('[ResultDownload] HTML preview loaded successfully')
      } catch (error) {
        console.error('[ResultDownload] Failed to load HTML preview:', error)
        alert('Failed to load HTML preview: ' + error.message)
      }
    }

    // Copy HTML to clipboard
    const copyHtmlToClipboard = async () => {
      try {
        await navigator.clipboard.writeText(htmlContent.value)
        copySuccess.value = true
        setTimeout(() => {
          copySuccess.value = false
        }, 2000)
      } catch (error) {
        console.error('[ResultDownload] Failed to copy to clipboard:', error)
        // Fallback: select textarea content
        if (htmlSourceTextarea.value) {
          htmlSourceTextarea.value.select()
          document.execCommand('copy')
          copySuccess.value = true
          setTimeout(() => {
            copySuccess.value = false
          }, 2000)
        }
      }
    }

    // Download with refresh check
    const downloadFileWithRefresh = async (index) => {
      // First refresh to get latest results
      await refreshResults()

      // Wait a bit to ensure backend has finalized
      await new Promise(resolve => setTimeout(resolve, 500))

      // Now download with updated results
      const url = `${API_BASE_URL}/api/download/${props.jobId}/${index}`
      console.log('[ResultDownload] Downloading file from:', url)
      console.log('[ResultDownload] Current results:', currentResults.value)
      console.log('[ResultDownload] File info:', currentResults.value[index])

      const link = document.createElement('a')
      link.href = url
      link.download = currentResults.value[index].filename || `download_${index}`
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
    }

    const downloadFile = (index) => {
      const url = `${API_BASE_URL}/api/download/${props.jobId}/${index}`
      console.log('[ResultDownload] Downloading file from:', url)
      console.log('[ResultDownload] File info:', currentResults.value[index])

      const link = document.createElement('a')
      link.href = url
      link.download = currentResults.value[index].filename || `download_${index}`
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
    }

    const downloadAll = () => {
      const url = `${API_BASE_URL}/api/download/${props.jobId}`
      console.log('[ResultDownload] Downloading all as ZIP from:', url)

      const link = document.createElement('a')
      link.href = url
      link.download = `results_${props.jobId}.zip`
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
    }

    // Auto-refresh on mount
    onMounted(() => {
      console.log('[ResultDownload] Component mounted, auto-refreshing results...')
      setTimeout(refreshResults, 1000) // Refresh after 1 second
    })

    console.log('[ResultDownload] Component setup with initial results:', props.results)

    return {
      downloadFile,
      downloadFileWithRefresh,
      downloadAll,
      hasErrors,
      formatFileSize,
      refreshResults,
      currentResults,
      // HTML Preview
      showPreview,
      showHtmlSource,
      htmlContent,
      copySuccess,
      previewFrame,
      htmlSourceTextarea,
      previewHtml,
      copyHtmlToClipboard
    }
  }
}
</script>