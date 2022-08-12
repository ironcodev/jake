<%
    class clsMime
        public function GetMime(ext)
            Dim result
		
            Select Case ext
                ' -------------- Picture Formats -----------
                Case "jpg","jpe","jpeg":
                    result = "image/jpeg"
                Case "bmp":
                    result = "image/bmp"
                Case "gif","ifm":
                    result = "image/gif"
                Case "png","pnz":
                    result = "image/png"
                Case "jfif":
                    result = "image/pipeg"
                Case "ico":
                    result = "image/x-icon"
                Case "tif","tiff":
                    result = "image/tiff"
                ' -------------- Audio Formats -----------
                Case "au","snd":
                    result = "audio/basic"
                Case "mid","rmi","midi","kar":
                    result = "audio/midi"
                Case "mp3","mpga","mp2":
                    result = "audio/mpeg"
                Case "aif","aiff","aifc":
                    result = "audio/x-aiff"
                Case "m3u":
                    result = "audio/x-mpegurl"
                Case "ra","ram":
                    result = "audio/x-pn-realaudio"
                Case "wav":
                    result = "audio/x-wav"
                ' -------------- Video Formats -----------
                Case "mp2","mpa","mpe","mpg","mpeg":
                    result = "video/mpeg"
                Case "mov","qt":
                    result = "video/quicktime"
                Case "lsf","lsx":
                    result = "video/x-la-asf"
                Case "asf","asr","asx":
                    result = "video/x-ms-asf"
                Case "avi":
                    result = "video/x-msvideo"
                Case "wmv":
                    result = "audio/x-ms-wmv"
                Case "wma":
                    result = "audio/x-ms-wma"
                Case "wm":
                    result = "video/x-ms-wm"
                Case "wmx":
                    result = "video/x-ms-wmx"
                Case "wvx":
                    result = "video/x-ms-wvx"
                Case "swf":
                    result = "application/x-shockwave-flash"
                ' -------------- Packed Formats -----------
                Case "zip":
                    result = "application/zip"
                Case "tgz":
                    result = "application/x-compressed"
                Case "tar":
                    result = "application/x-tar"
                Case "lzh":
                    result = "application/octet-stream"
                Case "lha":
                    result = "application/octet-stream"
                ' -------------- Text Formats -----------
                Case "htm","html","stm":
                    result = "text/html"
                Case "bas","c","h","txt","diz","id","cfg","ini","sql":
                    result = "text/plain"
                Case "sql":
                    result = "application/sql"
                ' -------------- Document & PDF Formats -----------
                Case "doc","dot":
                    result = "application/msword"
                Case "pdf":
                    result = "application/pdf"
                Case "rtf":
                    result = "application/rtf"
                Case "ppsx":
                    result = "application/vnd.openxmlformats-officedocument.presentationml.slideshow"
                ' -------------- Application Formats -----------
                Case "exe","bin","class","dms":
                    result = "application/octet-stream"
                Case Else:
                    result = "application/octet-stream"
            End Select
            
            GetMime = result
        end function
    end class

    dim MimeHelper

    set MimeHelper = new clsMime
%>