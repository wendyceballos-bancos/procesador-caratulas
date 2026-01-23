
import streamlit as st
import pandas as pd
import numpy as np
import base64
import io
import os
from datetime import datetime, date
import warnings
import re

warnings.filterwarnings('ignore')

# Configuración de la página
st.set_page_config(
    page_title="🏦 Vaciado de Carátulas Bancarias",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Archivo de mapeo Flex → Moneda embebido (actualizado)
ARCHIVO_MAPEO_MONEDAS_BASE64 = """UEsDBBQABgAIAAAAIQCREKFLmQEAAPgGAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC8lU1LAzEQhu+C/2HJVXZTFUSkWw9+HFVQwWtMpt3QfJGZavvvzWa1iKxtZYuXDbvJvO+Tmc1kfLm0pniDiNq7mh1XI1aAk15pN6vZ89Ntec4KJOGUMN5BzVaA7HJyeDB+WgXAIkU7rFlDFC44R9mAFVj5AC7NTH20gtJrnPEg5FzMgJ+MRmdcekfgqKRWg03G1zAVC0PFzTJ97kgiGGTFVbew9aqZCMFoKSiR8jenfriUnw5VisxrsNEBjxIG470O7czvBp9x9yk1USsoHkSkO2ETBl8a/u7j/NX7ebVZpIfST6dagvJyYVMGKgwRhMIGgKyp8lhZod0X9wb/vBh5Ho73DNLuLwtv4aBUb+D5ORwhy2wxRFoZwH2nPYtuc25EBPVIMZ2MvQN8197CEdJZ9Q55N+6QCYslLCWYqovYJC8XSN6+WMM1gX2IPuDwsq5FWz2IpGF9Kvv+7h6Gk8H1Hs5w+t8MqUXkAqRmGeHv5l/dsI0uw06ZXzumRjt4t9C2cgXqr95dpfaU7B5znu+tyQcAAAD//wMAUEsDBBQABgAIAAAAIQATXr5lAgEAAN8CAAALAAgCX3JlbHMvLnJlbHMgogQCKKAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAArJJNSwMxEIbvgv8hzL072yoi0mwvRehNZP0BMZn9YDeZkKS6/fdGQXShth56nK93nnmZ9Wayo3ijEHt2EpZFCYKcZtO7VsJL/bi4BxGTckaN7EjCgSJsquur9TONKuWh2PU+iqziooQuJf+AGHVHVsWCPblcaThYlXIYWvRKD6olXJXlHYbfGlDNNMXOSAg7cwOiPvi8+bw2N02vact6b8mlIyuQpkTOkFn4kNlC6vM1olahpSTBsH7K6YjK+yJjAx4nWv2f6O9r0VJSRiWFmgOd5vnsOAW0vKRFcxN/3JlGfOcwvDIPp1huL8mi9zGxPWPOV883Es7esvoAAAD//wMAUEsDBBQABgAIAAAAIQBpeUttFAQAALIJAAAPAAAAeGwvd29ya2Jvb2sueG1srFXbbuM2FHwv0H9QhbzKEnWzLMRe6GJhAyTbwHGT9imgJdpiI4lakkocBPtV/YT+WA9ly4mbonCzBWxJvA3nHM4cnn/a1pX2SLigrJnqaGTpGmlyVtBmM9V/WWZGoGtC4qbAFWvIVH8mQv80+/GH8yfGH1aMPWgA0IipXkrZhqYp8pLUWIxYSxoYWTNeYwlNvjFFywkuREmIrCvTtizfrDFt9B1CyE/BYOs1zUnK8q4mjdyBcFJhCfRFSVsxoNX5KXA15g9da+SsbgFiRSsqn3tQXavz8GLTMI5XFYS9RZ625fDz4Y8seNjDTjD0bqua5pwJtpYjgDZ3pN/FjywToaMUbN/n4DQk1+TkkaozPLDi/gdZ+Qcs/xUMWd+NhkBavVZCSN4H0bwDN1ufna9pRW530tVw237BtTqpStcqLOS8oJIUU30MTfZEjjp418YdrWDUnniOpZuzg5yvuVaQNe4quQQhD/Aw0bIdq58JwogqSXiDJUlYI0GH+7i+V3Ozc8BOSgYK1xbka0c5AWOBviBWeOI8xCtxjWWpdbzaZVCA5QoiWrLB3PG9kSgxJy2jzU55AnIgzKioaUOF5DgHhSzIBp64ss3BRkxovQO4pAUTpmWPNAgsBzfAgj//aCAj2gpDVRCmoxnazxK0bV7zjqyw0JbsK27MN2bA7533H+yAc3Ua5iERu++/JxzywcNB8teSa/B9kV7Csd/gRxABcnSt2BeJCzjm4P7FjmIn8p3McFDiGm4aB0Y8dlwD+SgYR/Yc6hH6BmFwP8wZ7mS5V5YCneouyOjd0BXeDiPICjtavBJ4SSxkRYETG1mcwXaRHRtB6lqGF/iJb0VjO8mCbypUVUNvKXkSrxpUTW17R5uCPU11A9ngnOfj5lM/eEcLWSoRWy5M2fV9JnRTAmPkjdU68JpiNtVf7HEcZ5k1NlLXDQzXSZAxSZPUsOd+7GWRh1yVAJX8N5T6ag3U+rfW9A77zH7HCG4FVchVcuGbh2oLflGgHmBYBU6iDSmUMQHjTWuPdL+tmnp0n1HlpxRLDIIiyq85rm4GeAiipEVB1PWkz/rNfzqLzlB4Fp8hxz8H/R92AfLHewJQDpZWr57qBFn2RHEkW3kpZP8GN1HID8Qfja2Ja1hzxzPcYGIbgevYRuKm9twbz9N57CmBqOsu/D+Kfm/qcLhHFUswr1yCSx/g9l2QdQzZUEGrMwG+b8nGXhBbDlB0M5QZLppYRhz7ruGlmeONUZrMveyVrAp//cGSG5j9aoJlB+VIVaK+Hapntu89dK53HfvjPbJ9uEhVIPvV/zbxBqKvyImTs9sTJyZfrpZXJ869nC/v77JTJ0dXcRqdPj9aLKLflvNfhy3Mf0zo7sDVs5epOchk9hcAAAD//wMAUEsDBBQABgAIAAAAIQAuZhB6PAEAAOEEAAAaAAgBeGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHMgogQBKKAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC8lE1rwzAMhu+D/Yfg+6Kk3doxmvYyBr1uHexqHOWDxnaw1G399xPt1rTQpZfQi0ESfvUg+9Vs8W2b6BMD1d5lKo0TFaEzPq9dman31cvdo4qItct14x1maoukFvPbm9krNprlElV1S5GoOMpUxdw+AZCp0GqKfYtOKoUPVrOEoYRWm7UuEUZJMoFwrKHmJ5rRMs9UWObSf7VtpfNlbV8UtcFnbzYWHZ9pAWZD7O2HbURUhxI5U3HcZaFmtONYkBWcpxkPSUO8bWScB5R93Nd+eu1hjPpoRkPSsHwZ7GaxC2F3pn0M6ZAMXz6sqULkjuOQIthVemEm136eXpqHf2hsbYInX3BsvIW9acQs6RTS5NSS0Mpi8K6bxj6m33zfu9wPapRKB8zfOMhWOvbLcfoPBk4W0/wHAAD//wMAUEsDBBQABgAIAAAAIQBMgXCbvAkAAMZIAAAYAAAAeGwvd29ya3NoZWV0cy9zaGVldDEueG1snJPbitswEIbvC30HoXvHlo+JibPsert0oRelx2tFHsciluVKyomy796xc9iFQAkLtnWw/u+fkUbzu71qyRaMlborKJsElEAndCW7VUF//njyppRYx7uKt7qDgh7A0rvFxw/znTZr2wA4goTOFrRxrs9934oGFLcT3UOHf2ptFHc4NCvf9gZ4NYpU64dBkPqKy44eCbm5haHrWgp41GKjoHNHiIGWO4zfNrK3Z5oSt+AUN+tN7wmtekQsZSvdYYRSokT+vOq04csW896zmAuyN/iE+EZnm3H+yklJYbTVtZsg2T/GfJ3+zJ/5XFxI1/nfhGGxb2ArhwN8RYXvC4klF1b4CoveCUsvsGG7TL6RVUH/Ziwos5R98uI4nXlxlATeQ5KVHpsGJYuzMmH34QtdzCuJJzxkRQzUBb1n+QOLMuov5mMF/ZKws2/6xPHld2hBOEAXRonT/ReoXQltO6jDlJKhZpdarwftM64K0MaOmsGGCye3cF4fIcL+OVnjAH39i/Hb/jmIp7HSvxqy5BZK3f6WlWswErxRFdR807pvevcZ5KpxOJvgjgwllVeHR7ACaxnDmYTJ4CN0i1D8EiWHS4m1yPdjuzsy42gSh0k2ZbieiI11Wp3dTvqjEs9wVGJ7Ukbpf5X+aP0PAAD//wAAAP//lJzbTtxIFEV/BfUHBJfvjgjSRPkRxCDlKYkCSmb+fuxqRHfty4T9EiFY9K4qVy+O23Vy9/z16enly8PLw/3dz++/b35+OpXTzfOPh2/P+1cf+9PN15f9i+nDMp1u/injw+PHv//98vT8+PRt/373oZ9O93ePx6/9dfxe/e39B8/7d3/dd3e3v+7vbh9fic9MlGF4Y273/LdB7MHXg/j/5AP+dNr/fUsukCyIYdTJQ5J8wG1yD8lMFJc8JskH3CZfVrJej89M2OT92r5/tQ+4Tb6s5DmZCZs8J8kH3CZPsNpM2OQlST7gNnmGZCbKcBlds7fXJPmA2+QFkpmwc96S5ANuk1dIZsLOuXRJdKXb7A1dcrwgvOndG6uA0v7gsbOmGp2QyQRj0yOXFaEqtJlibHrksyJ0hUZTjNvqJXJapfGa4nVXWjNvtBJ5rdKQjmZTjJ175LYi1IV2U4y97pHfitAXGk4xwwVqFFcix1UaVh4tpxg798hzRWgMTacYl95Hqqs0zB1dpxibHrmuZ4/16DrBFHfd+6xuY9f1VLmp0s3suj5yXaWheEPXCcbPPXJdzx7rsYITjE+PXLeX7VS4ousEYwuLPnJdpWHl0XWC8emR63p2XY+uE4xPj1zXc7nWo+sE46975LqeXdej6wRj04fIdZWG646uE4xd+SFyXaXhlgldJxg/98h1A3tsQNcJxs89u0/lum6gO1V1q3rZms3f9yFyXaVh5dF1gimDS49cN7Drrv54nu9YBePTI9cNXNddFYyv6eq21c09ct3Arrv60/2arm5dXXrkuoFdd3VJX9PV7atLj1w3sOsGdJ1g7HUfI9dVGvY8uk4wPj1y3ch13YiuE4xPj1w3sutGdJ1gfHrkupE9NqLrBOPTs8/luK4b6ZM5dQ9r9vwYua7S7a4bsa4TjP9IMnLdyB4bsa4TjF/5yHUje2zEuk4wPj1y3cgeG7GuE4xPj1w3sutGdJ1gbPoUua7SsOvQdYLx6ZHrJnbdhK4TjE+PXDex6yZ0nWB8euS6iV03oesE49Mj103ssQldJxifnj2H4LpuoicRwaOIKXJdpds9P6HrBOMfwUSum9h1E7pOMD49ct3ErpvQdYIpV+VXczcxRa6rNKw8uk4w/hFU5LqZnztM6DrB2LnPkesq3c59RtcJxqdHrpvZdTO6TjB+5SPXzey6GV0nGD/3yHUzu25G1wnGzz1y3cwem9F1gvHp2XNXrutmevKq7mEvb8vm/T5Hrqs07Hl0nWD83CPXzey6GV0nGL/rItfNXNfN6DrB+IfekesWdt2MrhOMT49ct3Bdt6DrBGNXfolcV+l21y3oOsH4uUeuW9h1C7pOMD49ct3CrlvQdYLxKx+57jhphEce0HWC8emR6xb22IKuE4xPz86ZcF230EkT9XmdMe0Sua7SsOfRdYLxuy5y3cKuW9B1grHpa+S6SsPc0XWC8emR61Z23YquE4xPj1y3cl23ousE49Mj163suhVdJxifHrluZdet6DrB+PTIdSu7bkXXCcanR65b2XUruk4wPj1y3coeW9F1gvHp2bk6rutWOlkXHK1bI9dVurXNiq4TjJ37Frmu0pCOrhNMGS5QU89vkesq3aZv6DrB+PTIdRu7bkPXCcavfOS6jV23oesE49Mj123sug1dJxifHrluY9dt6DrB+PTIdRu7bkPXCcbvush1G7tuQ9cJxs89ct3GHtvQdYLxc49ct3Fdt9FJYnEGz57u68KzxHwXu/Fp4uQ4cZedJ654K7zS0YliRfkliJxXOnGIrqNTxYpyyi9d5L0zjmtAJ4vriwLl1yByX+nEQ9cO7ScpvwaR/0onHkZ0dMJYUVfP7ttjtl3kwFJxvAp0ylhRfgSRB0snbl47OmmsKD+CyIWlE0VdR6eNFeVHEPmwdEJ2HZ04VpQdQdheIfor9lHhSXtF+RFkTqwtGbATC3dZiDYLP4LMibLRgpwYtVrUzoir1rk/9ZmIA3eFnKjaLa5O07Q+CPstVMNFIScqyo8gc6JqqCjkRNmacblWsAaZE1VTRSEnRq0XtVMi2AfCiYWcqNov/FXInKiaKwo5UVF+BJkTVYNFIScmbRgl68M442gkcqJqxbBrUDsn3r8PVKMFdWMU2Y5hellL1o9xxmENqCNDUr75LasTVcNFT05UlL8KWZ2omi6oM6PI1gx7FTInqsaLnpyYtGeUrD/jjOM+ICcmLRqldlQE7wXhROrSOL8o3i+YD6lK7aoIRiDqROrUOL/ou0eQOVE1YlC3RpHtGm4Nsn6NopoxenKibNmwI8jqRNmQQXWipNy7sXZZvH8fyKYMqhMVZY00ZE6sOO4x7soV1aRvCs6cqJoz6Ez1/mFV0JJdOy6CqyDunelc9d59ziO4OoHc1om1MyMYgTjgQmer92ahZATZvbNo5ih0wnk/XCmuAj4evr38tx//AQAA//8AAAD//zTOQQ6CMBQE0Ks0f19ppaVKgAQa3XmIGj/QiJZ8SjQx3l3UsJs3i8kUbo7h6IeIxAjbEmqZNzLNgD0pn/2lhJeRwppMHrhS2Z6rVAveaGO53AkrlbFa1ts3JFUxug5Pjjp/n9iAbSxBbAww8l2/5hjGX6uBnUOM4baqR3dB+ioF1oaw/Plj2U0ega5TjxirDwAAAP//AwBQSwMEFAAGAAgAAAAhAJMJR0DBBwAAEyIAABMAAAB4bC90aGVtZS90aGVtZTEueG1s7FrNjxu3Fb8HyP9AzF3WzOh7YTnQpzf27nrhlV3kSEmUhl7OcEBSuysUAQrn1EuBAmnRS4HeeiiKBmiABrnkjzFgI03/iDxyRprhioq9/kCSYncvM9TvPf7mvcf33pBz95OrmKELIiTlSdcL7vgeIsmMz2my7HpPJuNK20NS4WSOGU9I11sT6X1y7+OP7uIDFZGYIJBP5AHuepFS6UG1KmcwjOUdnpIEfltwEWMFt2JZnQt8CXpjVg19v1mNMU08lOAY1E5ABs0JerRY0Bnx7m3UjxjMkSipB2ZMnGnlJJcpYefngUbItRwwgS4w63ow05xfTsiV8hDDUsEPXc83f1713t0qPsiFmNojW5Ibm79cLheYn4dmTrGcbif1R2G7Hmz1GwBTu7hRW/9v9RkAns3gSTMuZZ1Bo+m3wxxbAmWXDt2dVlCz8SX9tR3OQafZD+uWfgPK9Nd3n3HcGQ0bFt6AMnxjB9/zw36nZuENKMM3d/D1Ua8Vjiy8AUWMJue76Gar3W7m6C1kwdmhE95pNv3WMIcXKIiGbXTpKRY8UftiLcbPuBgDQAMZVjRBap2SBZ5BHPdSxSUaUpkyvPZQihMuYdgPgwBCr+6H239jcXxAcEla8wImcmdI80FyJmiqut4D0OqVIC+/+ebF869fPP/Piy++ePH8X+iILiOVqbLkDnGyLMv98Pc//u+vv0P//ffffvjyT268LONf/fP3r7797qfUw1IrTPHyz1+9+vqrl3/5w/f/+NKhvSfwtAyf0JhIdEIu0WMewwMaU9j8yVTcTGISYWpJ4Ah0O1SPVGQBT9aYuXB9YpvwqYAs4wLeXz2zuJ5FYqWoY+aHUWwBjzlnfS6cBnio5ypZeLJKlu7JxaqMe4zxhWvuAU4sB49WKaRX6lI5iIhF85ThROElSYhC+jd+Tojj6T6j1LLrMZ0JLvlCoc8o6mPqNMmETq1AKoQOaQx+WbsIgqst2xw/RX3OXE89JBc2EpYFZg7yE8IsM97HK4Vjl8oJjlnZ4EdYRS6SZ2sxK+NGUoGnl4RxNJoTKV0yjwQ8b8npDzEkNqfbj9k6tpFC0XOXziPMeRk55OeDCMepkzNNojL2U3kOIYrRKVcu+DG3V4i+Bz/gZK+7n1Jiufv1ieAJJLgypSJA9C8r4fDlfcLt9bhmC0xcWaYnYiu79gR1Rkd/tbRC+4gQhi/xnBD05FMHgz5PLZsXpB9EkFUOiSuwHmA7VvV9QiRBpq/ZTZFHVFohe0aWfA+f4/W1xLPGSYzFPs0n4HUrdKcCFqPjOR+x2XkZeEKhAYR4cRrlkQQdpeAe7dN6GmGrdul76Y7XtbD89yZrDNbls5uuS5AhN5aBxP7GtplgZk1QBMwEU3TkSrcgYrm/ENF11YitnHILe9EWboDGyOp3Ypq8rvk5wULwy5+n9/lgXY9b8bv0O/vyyuG1Lmcf7lfY2wzxKjklUE52E9dta3Pb2nj/963NvrV829DcNjS3DY3rFeyDNDRFDwPtTbHVYzZ+4r37PgvK2JlaM3IkzdaPhNea+RgGzZ6U2Zjc7gOmEVzq54EJLNxSYCODBFe/oSo6i3AK+0OB2fFcylz1UqKUS9g2MsNmR5Vc0202n1bxMZ9n251mf8nPTCixKsb9Bmw8ZeOwVaUydLOVD2p+G+qG7dJstW4IaNmbkChNZpOoOUi0NoOvIaF3zt4Pi46DRVur37hqxxRAbesVeO9G8Lbe9Rr1jBHsyEGPPtd+yly98a52znv19D5jsnIEwNbirqc7muvex9NPl4XaG3jaImGckoWVTcL4yjR4MoK34Tw6y/vuPxVwN/V1p3CpRU+bYrMaChqt9ofwtU4i13IDS8qZgiXoEtZ4CIvOQzOcdr0F7BvDZZxC8Ej97oXZEo5fZkpkK/5tUksqpBpiGWUWN1kn809MFRGI0bjr6effhgNLTBLJyHVg6f5SyYV6wf3SyIHXbS+TxYLMVNnvpRFt6ewWUnyWLJy/GvG3B2tJvgJ3n0XzSzRlK/EYQ4g1WoH27pxKOD4IMlfPKZyHbTNZEX/XKlOe/a1DriIfY5ZGOC8p5WyewU1B2dIxd1sblO7yZwaD7ppwutQV9p3L7utrtbZcUR87RdG00ooum+5s+uGqfIlVUUUtVlnuvp5zO5tkB4HqLBPvXvtL1IrJLGqa8W4e1kk7H7WpvceOoFR9mnvsti0STku8bekHuetRqyvEprE0gW+Ozstn23z6DJLHEE4RVyw77WYJ3JnWMj0VxrdTPl/nl0xmiSbzuW5Ks1T+mCwQnV91vdDVOeaHx3k3wBJAm54XVthW0Nnt2YK62OWi2YLdCmdt7LV+1RbeSmyOWbfCZmvRRVtdbU7Uda9uZtYOy57apGFjKbjatSIc/wsMrXN2mJvlXsgzVyrvtOEKrQTter/1G736IGwMKn67MarUa3W/0m70apVeo1ELRo3AH/bDz4GeiuKgkX37MIbTILbOv4Aw4ztfQcSbA687Mx5Xufm6oWq8b76CCELrK4jsiwY00R85eOBIoBWOgnrYCweVwTBoVurhsFlpt2q9yiBsDsMeFO3muPe5hy4MOOgPh+NxI6w0B4Cr+71GpdevDSrN9qgfjoNRfegDOC8/V/AWo3Nubgu4NLzu/QgAAP//AwBQSwMEFAAGAAgAAAAhAEqMQeoqAwAAKggAAA0AAAB4bC9zdHlsZXMueG1spFVbb9MwFH5H4j9Yfs9yaVPaKgla20VCAoS0IfHqJk5q4UtkuyMF8d85zqXNNAZjPMU+l8+fz+dzkrxtBUf3VBumZIrDqwAjKgtVMlmn+PNd7i0xMpbIknAlaYpP1OC32etXibEnTm8PlFoEENKk+GBts/Z9UxyoIOZKNVSCp1JaEAtbXfum0ZSUxiUJ7kdBsPAFYRL3CGtRPAdEEP312HiFEg2xbM84s6cOCyNRrN/VUmmy50C1DeekQG240BFq9XhIZ310jmCFVkZV9gpwfVVVrKCP6a78lU+KCxIgvwwpjP0genD3Vr8Qae5res+cfDhLKiWtQYU6SpviGRB1JVh/leqbzJ0LFB6issR8R/eEgyXEfpYUiiuNLEgHlesskgjaR1w3Vhn0kWitvrnYigjGT70vcoZO8iFYMBDAGX1HpqeUJXsXNR7Y5fzpwC3hbK/ZI5QzQvA3yo8ROjoG+DDOJyXqDVkCb8lSLXPwomF9d2qgFhKefU8EXH+NrjU5hVH8/ASjOCudJvW2U0DX+xTn+Xa1u86XDmb/lMOfUHbV7uh1H7jlXukSmnp8Ck713pQlnFYWcDWrD+5rVeNOUdbCw8+SkpFaScKdgGPGsADYgnJ+6xr/S/UAu62QPIpc2HdlimGEOOnHJfAalj1ev3H4U7QeewI7A8r/Dova6oz/VHYI/AZSEUZTUudsRJqGn1zLuGYYdpBz2V1zVktB+4AsgQfbb9FBafYdEl1rFeCnfTO01dPXARYjIajdMwiNxYNyTTR5oMi5tsj1cYo/ujHMYSIM9UH7I+OWyd+oAZhle9G36zXrRmqn/PkUoFrSihy5vTs7U3xZf6AlOwq42xD1id0r20Gk+LJ+755huHAvnbb2vYGJAV901CzFP242b1a7mzzylsFm6c1nNPZW8WbnxfPtZrfLV0EUbH9OBvt/jPXuPwQShfO14TD89XDZgfztxZbiyaan37U70J5yX0WL4DoOAy+fBaE3X5Clt1zMYi+Pw2i3mG9u4jyecI9fOP4DPwz7H4kjH68tE5QzOWo1KjS1gkiw/cMl/FEJ//KTz34BAAD//wMAUEsDBBQABgAIAAAAIQAmVXdY7QMAAPoiAAAUAAAAeGwvc2hhcmVkU3RyaW5ncy54bWykmttO20AQhu8r9R0i39feXZ9ClQS1FK6gRdBI7WWUGIiUOBQbRN++86/jYhJwtPsjBLGtL3PYmdnZSUbHz+vV4Kl4qJabchzoUAWDopxvFsvydhxMf559GgaDqp6Vi9lqUxbj4G9RBceTjx9GVVUPhC2rcXBX1/efo6ia3xXrWRVu7otSntxsHtazWi4fbqPq/qGYLaq7oqjXq8golUXr2bIMBvPNY1mPA5MnweCxXP55LE6aOzrJg8moWk5G9eRsVTwPTm+Keb182oyiejKK8KB5qHUcaq2NDnWs8DIOlfy0f+R/87u91YsbDhcVGOm5B55a20XvVB/Jy9TR9gYXvVPlj1vpUJ6QPvTGrfSMwknbFcKGsB1hQ+CIeX+cCBt43ih36QqLjYRVQ2StXDol7BY3gmtZdwJXEv4MLnZ44rBd+StvcVl3Rrq/5yFdy7ozrpNqw+Bk2EjkEq6TbGNwpAwT85z03N92ZJwXnrT5nigsXOKY7x0cyhM4os4fN0hYQjoWzh/XnOs06TpZBEZ5lErCds51ttISCye9jafykjIJ2gPGdi5oYy5sjvptj2Xn3uuoxdjt/p7EnO1xf8alSule6f3KH8T7o+4g3p8yh/C8P2Xextt+XqW20jq2heiHmsYsxUlKLt0asw4O291wuzdY6Ym2yr9/jnsr6lrchBSugSNo/aRbHK4jlPfHrXR4npDO4TiLM65DxhHKu+MQt435GHWrRzpyYb/aNLjUeTt8YHAUasfZRUd5jVrnh1vlUWn9cZwiGeU523F+Z6T7r7scheIDxepg2PQH7UEc24TjwhnMenAAN0eodXLpVOe7uIQ/g8u6M7i4jsEl5l1xtJLNFpnZY2DmuEW+4LYp9cdtvjvicVus4hjrLpdO697FMW31w3VISbe4v3QD6eIGR+UzJLld9wx9jTQgTq57hWcULuc4X+kmFOVzTvmYky6rR7hOhk6U50nbh5ztXNhIZ0Wtu7vyGtPVptblGiNbx2rTwTE5YXDsV4R0zOcJHFWTwHEq8sdt0fXDpa/L7WcTrrtMW+uUjtEaSZPj8oHa/1KpmnOcP574Sxfbtd1hCek4yxA4Omo/HMrbcZ+/dHsQI6Rznieki85ae0g37SehJrU9rePsooujKSVwUd4Tl4bczm0YnJSOft7Pdqu8bBMMLnXeEU9ejsBaDqG4fK9YvTU1ypBmssdJa2QExaVLrbuQr0QsZrtDgen1t91bl6ffd2+dnF/u3vp6db576+LXHjj9Pd17rx977/Xl6np/WNEZ0OHU6bbQra80fIVe0s1Xr3B5D0f8dHr13vAFh3CnAUQk312Z/AMAAP//AwBQSwMEFAAGAAgAAAAhAK0x1F2kAAAA2gAAABUAAAB4bC9wZXJzb25zL3BlcnNvbi54bWxkzb0OwjAMBOAdiXeovJO0DKiq+rMxMcIDRKnbRGrsKrZQeXuKGLue7r5rhy0txRuzRKYOKlNCgeR5jDR38HreLzUUoo5GtzBhBx8UGPrzqV33DdMjihY7QdJBUF0ba8UHTE5Mij6z8KTGc7I8TdGjlTWjGyUgalrstaxqq+EX4bi3EpIK/L1mO4i8Iu1fE+fkVAzn+eCVN5tcJLD9FwAA//8DAFBLAwQUAAYACAAAACEAkDCaAz8BAABLAgAAEwAoAGN1c3RvbVhtbC9pdGVtMS54bWwgoiQAKKAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAArJKxboMwFEV/JfJujIEARkBUZW2kSu3Q9WE/J5bARrbT5PNL0qTtUKkdur3lnnN19drNeRpXb+iDcbYjPEnJCq10yth9R45R05ps+nZuZu9m9NFgWC0JG5q5I4cY54axIA84QUgmI70LTsdEuok5rY1ElqVpySaMoCAC+6KQG+YczCfodDolpzxxfn+Jcfa6e3y+sqmxIYKVeE/N8m92Y7WbIR4uvIo9gY8W/dbZ6N0YSN8qJ48T2rgDC3u8XH07Sl2Vmq/XElWhlByKVFQ8L7TOZZ5n+qN4RypZDbXQNR30kNJCl4oCiJxCpXQGHMo644viBf102+x/OrMrsW/Zb0UXN5y3EOXhYRzvrUWRwyByTnkFJS2wFnQAxSmiEmmmdVELsawcTGPN2JHoj0jYIvtpKfb9Lfp3AAAA//8DAFBLAwQUAAYACAAAACEABSIHPD4BAAAjAgAAGAAoAGN1c3RvbVhtbC9pdGVtUHJvcHMxLnhtbCCiJAAooCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACkkcFrwyAYxe+D/Q/BuzVJk6ilacmWFnobY4Ndv0RthahB7RiM/e9L6C7d6GkneX6833ufrrcfZkjepQ/a2RplixQl0vZOaHus0evLHjOUhAhWwOCsrJF1aLu5v1uLsBIQIUTn5SFKk0wXejoPbY0+acN21QPPccsbiotdU2Je8hTzptwzSouyrXZfKJmi7YQJNTrFOK4ICf1JGggLN0o7DZXzBuIk/ZE4pXQvW9efjbSR5Glakf48xZs3M6DN3OfifpYqXMu52tnrPylG994Fp+Kid+Yn4AI2MsK8HRn9VMVHLQMi/4Bqq9wI8TTTKXkCH630j85G74bbZNrTjnHFcKe6FBeqEhiALzFQoXLIoGJ5drMWL5bQ8WWGMwoVLiTjuAORYSkFT3OlCsb5bCa/Hm7WVx+7+QYAAP//AwBQSwMEFAAGAAgAAAAhAL2EYiOQAAAA2wAAABMAKABjdXN0b21YbWwvaXRlbTIueG1sIKIkACigIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGyOOw7CMBAFr4LSky3o0OI0gQpR5QLGOIqlrNfyLh/fHgdBgZR6nmYediS8dRzVRx1K8p3BE2caPKXZqpfNi+Yoh2ZSTXsAcZMnKy0Fl1l41NYxgUw2+8QhKjx28LVptcFYXdIY7INUXzE9uzvV1Dlcs81lSSH8IB5vQdcnH4IX/1zHC0D4O27eAAAA//8DAFBLAwQUAAYACAAAACEAPShKdfAAAABPAQAAGAAoAGN1c3RvbVhtbC9pdGVtUHJvcHMyLnhtbCCiJAAooCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABkkMFqwzAMhu+DvkPwPbGTlnQpSQqtU+h1bLCrcZTGEFvBcsrG2LvPYaduJ/FJSN+P6uOHnZI7eDLoGpZngiXgNPbG3Rr29npJn1lCQbleTeigYQ7Zsd081T0dehUUBfRwDWCT2DCxXmXDvrpcFGVVdmkhZZXuTrtzWpUrXuQ5F/v99tTJb5ZEtYtnqGFjCPOBc9IjWEUZzuDicEBvVYjobxyHwWiQqBcLLvBCiJLrJertu51Yu+b53X6BgR5xjbZ4889ijfZIOIRMo+U0Kg8zmnj8vuUaXYie8DkDX2MQ423N/0hWfnhC+wMAAP//AwBQSwMEFAAGAAgAAAAhAKy37IvgCgAA+DkAABMAKABjdXN0b21YbWwvaXRlbTMueG1sIKIkACigIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOxb227jyBF9X2D/gWBekgdapKgbhdEsbNmzMTKzM1hrN3lbNJtNizHF5rCbvmCRT8pX5MdSzeb9IpGUN/EGmQFmLIpVXVVdXVVdp/zuu+eDrzySiHk02KjGha4qJMDU8YL7jRpzV1up371/h/ka04CTgO9eQnKH9+SAFHj4y0ZVlQPK/y+99AM6kI16TXF8ADKavFb6+vZ6o+rPugF/dXN7eXW10heWeTM1VrOrq8XyZnptXRmXN1emPr2q0/6ci7uqf3VNGI68kCfabCOCIiWIySNVnEyQizrJHaYhSJo8Tg0hhMOu5U5XU2NuLKbmbIUtrDvInpn2HLlLe+WqClguYGvMN+qe83A9mbDELuzi4OGIMuryC0wPE+q6HiaTqa4vJgfCkYM4mpQskTE6oDGMwgikj7hHWML8kvPIs2NOmPr+22/ePTNnLaVSOIruCRe7wkKEQeHhQhdrJcaKKAXdeRST5KPrEd9hwnQzY+oac2QsTcNcuMul5VoGsla6aZnTlbWyVCVgU+kzATPlD9KYIG8u2NPT08WTeUGje2E7Y/K3Tx+l42UGe2b93w3P1VfKB3JvVGtmItsyDc1YooU2IytLs5FjaIQ4lj513dnKAhUzAnOjLjE4jOWuNNu1dW3mLhwNIcvU0NJxp8hAC3CzfLu8Q0gjrgTFRvVab5Jtd5O+1/I5PfGJOLCJABu1tOXZAuDToU+eRSDIXYx8jSFq5J+rPLKj9wkF6D5hnivbwgv5fsY2YxMRd6MKl7nbo4g4f/X4/icGMQDczgs+YxxH4Am62lChhe4aDqDn96U015+I46E7Ej3CEf6UHt6ey1aJPyDGz2JwGXP6F/LyhXoBHyf/edTXiJMdeiDBMPU/kuCe72+DOwIhzxknuFB9h+7HEX/e/jhqw74nAYmQSCQ77yAiXB9Xq+75zSMcpD8jtt9SZxyHjxQnIvRe3sfucgHBd46JM3McbM90a2mYM9c1sWlOIWv1PDI79LxFHO8vfX+U7p/tvxPM4bjBvzRKM/a4HbyDLI73X/JE1ybPRGS6NHQkP9ciS/IsjScivCSfWSls9SdKkvyp5NovR6QCfaDR4Zq4KPYhn36Nke9BLnWKNPcb5UTnUCTQ06VLM4pPOKQA2FKZ6ULcL8V6gUtDxPciqS8nX1DE4ZxtobiMKMTl7izWu1zpFPRIiuzD/LjgHfmzmbHQ2gsc8rxRoXYNPN9Htg/FWF5FOR4LffQiq+ctpFowkOdQBaJnUmVB9CdRgHz5Qht7KHqdz4H/kjI9krbTMyKMD4Eqz7rP8ElcBxQbMRAN3GQt0u0ncE+vqCqPpnzx/i1sdO2owlXhOcvXcWDTGCzhHCsGyidUnPBqYXFd2EpVhDNu1KTghRI4uD8VtUSBtVEvMQYp+C2cNskg01Y8qYTK+mZ1bHjKUFRH/WQ6IxLVwxfJdi6PcM0d7hPnmpYu/KwooHJHFvX8CUcWVL5PmOIQJWbCmzPHPurU5cU63JpBveuTcjEaEQYOgEXmLnnwD5ST3NPADWVdojwiPwY3mM7nsnxMzFPiUOSKyjr1dNK0WDV/5raamqeNBaQ0oIcXJUnACmRgZUv9+CADwN5zHAJX9Txi+B6DtPHrylwQy7GnmrFYYG22MJFmm4auGfpUX1kmXF/12T9arF0XlO3p0wdxlduoWQEApR9KKJ+I3fP68ypRJ4k42z2F2vsjpQ9x2DP6/Cw2tXygU+r28NPjmLef0dc6bpJ7r7Ki11VO+X9Z8b9eVnRcS/MwY0Ab71RM7uTRiDDVmqOT7jUidJJV5YEYGHDLYtUu24VZjGFmafAZYJoG7dsxT72VUNhnOsw+TUYDDNQkfjsWKrdKCuv0yN11wnKxXvbQ6nu/s9qmrEi5LVRYajbMj6pMBvhQlfBMM+4goZeqwGplebriS6zSbHcVRpmfNkpalf6RyW7Zn1rqtc5lztT+p+AhoE/BuQZIe6Wlll1hgEWf2jdp8wmarqNTYf0GtjzVWPYZC2WXp5W9eeYRwpw4SuJ6AkjrVjvj/zuOFY2WamGtHs2QctRp4TQgarRQvx0/qjeNCxv1uGeXbdRkNMBETeK3Y6FSU7y4V/coeGt0XfGl8tp/V+se3fzCAu21G0+bCaWGQTXG9Foj45I0BmTrsexqtwcA9lKEBq0rLcwb7gE2yFHS+vHEe7KNURjWRT4rIci30Hj4dZ7CGJrAMTQBZGgCydDKUIZsZmSSJb2CkpKMhYKT5ZqGhWdzzZ1bgNUu56Zm6dAkWS51x9SdJYCxAhNBa/CGwx2BRiAgidbKmGHH1OaLFdFcYumaeKKtnMXUcN2Va6+gvgEaFOA9jQSJa6P5zLVNDZvEAJKVAeDucqFhNLMXlgHjAzPRT0RrAOvLvRuPQSH4RCPBIjHDCJA1QTlDvN6BBnWApdzxAIhZwh7nwiDNZlfZFzqBn8JTB1aHRzg65XGPxMAtMQ667QQMnDfMjuTYLtAKrd9O+GuBwwrT9qgxy5vVymtAmmil/09e4wb0zmoYjziJgFy5gL0hzpIxE5iLeYAA1RjViYhWHn050V5D2QhOBdaLgxKwZ/sUP+SI3x+gNZ6iaA0IbbyUowZqvNMTNZoXMA6RD0K2XMIpgL8wjvzEkg6epFZiE+PCmBTvijhbwI5lguSb/E0KMN4JQDE7xxNqFyBSc/wl3fcjsknwM8v7+apObPsezMJFJNEphUgnoDWbfAUNAcA0J/psok8nDr6AqFhMoPSSIlP4NZZPeFVlqGJlQo0yhC7Rru3ul9oXOaxVAtDTWaDmy1m26prZcfAaQzSAAYAjQKDRMbjjwM0IfEVyEHB4FY4rQYndHNYeYCVcwOkDJZDoYGVK78jybeBjHpOhMoXEXymKdl4IE4lEgLqAtQK822mCNfe4X58dqWZ1eW1MIGVZllSX+tc/eewfWYDFSdIbZ95S/h3GQNr3QRY/x6uWpnHbB73A2XwU3McQwcfIAoM25J5GL0dpe8giNUsHSV+HWUQePQHID+SWn84goDwZI8qeZCMS2UOl489u7zEJkSowtOAJEzGF7wnM1B5sEinUVRh6hGc0UjIh2YWygzdQGPqCQKCwwAQg1ZDCVAEMOyiQcpU4hDFYQOmBW74EcuHAKwThfc7s4ttv2kSTdWxdC/kUVbQ93bbzxWAezDuLkZurs3c/KcazyHVI2Y5yRzmUcQcbF486H70momTqkW5x1qDw8eGYU4hkj5LpNYeM0sKl1/xwaptqMv0Cw6dwHNPvGkOs6+rYSj61kt3UKnMrlbSS3NCy92tLZ9SlkZWxtHI65Tg1+E9FD/G5OVTXIasMgq3zOWsxnQHjOcdU7ZjNERINJK7O4fQnv7reXjJGsQchz7mBEoK/jN5u4JVy6C5hStNQIoyVzVrfCPg2Py8QKGE0Gi7/coF8uj+1bv5e6krH6HqT3L0wTg636S1ALNmbNDMpJIUuul5uVughPaWufZ6mqxvebY46m2J0axiHFtuMlKVuqp5sqnGq8LyxoUpa9ryAJXlkDvMjcUkkflmglhX7hL6UkwMw/uDQl9FCm3Q0LQDAo2lFbV4eIR+iL3R1fvNg3bLT2Uhl/7DZudWjWYm9Hk8Mmz2eGHZ7PDFs93hi2O/hxLL/O/agC+r+E7unctSIYiEX4BwNkqwy/Jwla9c7DI0D+goVkVhIyjhuf4s55SO1UCJnPk44afvdzff/BgAA//8DAFBLAwQUAAYACAAAACEAmpRQpbsBAAB9BAAAGAAoAGN1c3RvbVhtbC9pdGVtUHJvcHMzLnhtbCCiJAAooCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC0lF1r2zAUhu8H+w9B94r8hT9KnZIlDhRWGF0HvT2WjhIzSzKSvGyM/ffJbsfompKNbVfmWJz3Pef1I19efVb94hNa1xldk3gZkQVqbkSn9zX5cLejJVk4D1pAbzTWRBtytXr96lK4CwEenDcWrz2qRXjRhef1tiZfi12TVuvthjZFldFs3cS0TDYZLcpmm+dvynyXNd/IIljrIONqcvB+uGDM8QMqcEszoA6H0lgFPpR2z4yUHcet4aNC7VkSRTnjY7BX96onq2meh+5blO5pOY022u6Zi+q4Nc5Iv+RGPRo8CCv0MG3HuNE+2N19GZCwf6Y62LCg9R06NjmtvbddO3p05zyOx+PymM55hABidn/z9v0c2X8Z7kXRKkuhrdKYxgXkNMOyoi2ImCKKKkqkzMqqerG54EVbVrKkrWwjmslcUIAqpVAImUAMeZnEf7+OeATlBjTscUbGh494NuEfBJ5ko9PSDOAPEyQFewfWa7SbgIg1/W8rn2B7AP4xTPmMPYv0JyrnMhlG289kCM6wn1d2LF7G7E8aPVrlznacDqkLV8Vq6JlpxeTJfrmSU/3kl7H6DgAA//8DAFBLAwQUAAYACAAAACEAOEpF7GMBAAB+AgAAEQAIAWRvY1Byb3BzL2NvcmUueG1sIKIEASigAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjJLfS8MwEMffBf+Hkvc2/aFDStuBk+HDBoKViW9nctuKTVqSzK7+9abtVjv0QchDkvve5753STI/itL5RKWLSqYk8HzioGQVL+QuJS/50r0jjjYgOZSVxJS0qMk8u75KWB2zSuGTqmpUpkDtWJLUMatTsjemjinVbI8CtGcV0ga3lRJg7FHtaA3sA3ZIQ9+fUYEGOBigHdCtRyI5ITkbkfVBlT2AM4olCpRG08AL6I/WoBL6z4Q+MlGKwrS17elkd8rmbAiO6qMuRmHTNF4T9Tas/4C+rlfPfatuIbtZMSRZwlnMFIKpVLZByVsnBwEKnAW+Q1lWOqETRTfNErRZ28FvC+T3bbYGVYCzOlgIs5tHVNK+An4l9LfWFut7Gyoid6zbeOjtHNlEi4d8SbLQD2euH7jBLA/u4tCu27fOykV+5364ECdD/yCGQR76cRTGN9GEeAZkve/LH5N9AwAA//8DAFBLAwQUAAYACAAAACEAJwHspZEBAAAXAwAAEAAIAWRvY1Byb3BzL2FwcC54bWwgogQBKKAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACcksFu2zAMhu8D9g6G7o2cbiiGQFZRpBty2LAASXvnZDrRJkuCxBjJ3mbPshcbbaOJs+3UG8mf+vWJoro/tq7oMGUbfCXms1IU6E2ord9V4mn76eaDKDKBr8EFj5U4YRb3+u0btU4hYiKLuWALnyuxJ4oLKbPZYwt5xrJnpQmpBeI07WRoGmvwMZhDi57kbVneSTwS+hrrm3g2FKPjoqPXmtbB9Hz5eXuKDKzVQ4zOGiB+pf5iTQo5NFR8PBp0Sk5FxXQbNIdk6aRLJaep2hhwuGRj3YDLqOSloFYI/dDWYFPWqqNFh4ZCKrL9yWO7FcU3yNjjVKKDZMETY/VtYzLELmZKehW+Qy5qLMzvX84cXFCS+0ZtCKdHprF9r+dDAwfXjb3ByMPCNenWksP8tVlDov+Az6fgA8OIfUEdr5ziDQ/ni/6yXoY2gj+xcI4+W/8jP8VteATCl6FeF9VmDwlr/ofz0M8FteJ5JtebLPfgd1i/9Pwr9CvwPO65nt/Nyncl/+6kpuRlo/UfAAAA//8DAFBLAwQUAAYACAAAACEAWUneyiwBAAASAgAAEwAIAWRvY1Byb3BzL2N1c3RvbS54bWwgogQBKKAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACkkcFOwzAMQO9I/EOUexc3ZaOd2k5L20k7ICExuEdt2lVakirJyibEv5MKxjhwgqNl+/nZTlcneUCjMLbXKsPhDDASqtZNr7oMP+82QYyRdVw1/KCVyPBZWLzKb2/SR6MHYVwvLPIIZTO8d25YEmLrvZDcznxa+UyrjeTOh6Yjum37WpS6PkqhHKEAC1IfrdMyGL5x+JO3HN1fkY2uJzv7sjsPXjdPv+Bn1ErXNxl+K+dFWc5hHtAqKYIQQhYkUXIfQAxAGS02ybp6x2iYiilGiku/eqGV89oTdNt46uiWh+HVOpPDCTwDICrWjMWwSKKKhvEdY4v7ipYJC9cVizw4JdeelFys/ukXXfweRNPzJ2FGf+Ot5J3Y8W7a/ufM3+eT6zPzDwAAAP//AwBQSwMEFAAGAAgAAAAhAHQ/OXrCAAAAKAEAAB4ACAFjdXN0b21YbWwvX3JlbHMvaXRlbTEueG1sLnJlbHMgogQBKKAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACEz8GKAjEMBuC74DuU3J3OeBCR6XhZFryJuOC1dDIzxWlTmij69hZPKyzsMQn5/qTdP8Ks7pjZUzTQVDUojI56H0cDP+fv1RYUi429nSmigScy7Lvloj3hbKUs8eQTq6JENjCJpJ3W7CYMlitKGMtkoByslDKPOll3tSPqdV1vdP5tQPdhqkNvIB/6BtT5mUry/zYNg3f4Re4WMMofEdrdWChcwnzMlLjINo8oBrxgeLeaqtwLumv1x3/dCwAA//8DAFBLAwQUAAYACAAAACEAXJYnIsMAAAAoAQAAHgAIAWN1c3RvbVhtbC9fcmVscy9pdGVtMi54bWwucmVscyCiBAEooAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAITPwWrDMAwG4Huh72B0X5z2MEqJ00sZ5DZGC70aR0lMY8tYSmnffqanFgY7SkLfLzWHe5jVDTN7igY2VQ0Ko6Pex9HA+fT1sQPFYmNvZ4po4IEMh3a9an5wtlKWePKJVVEiG5hE0l5rdhMGyxUljGUyUA5WSplHnay72hH1tq4/dX41oH0zVdcbyF2/AXV6pJL8v03D4B0eyS0Bo/wRod3CQuES5u9MiYts84hiwAuGZ2tblXtBt41++6/9BQAA//8DAFBLAwQUAAYACAAAACEAe/MCo8MAAAAoAQAAHgAIAWN1c3RvbVhtbC9fcmVscy9pdGVtMy54bWwucmVscyCiBAEooAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAITPwWrDMAwG4Hth72B0X5x0MEqJ08so5DZGB7saR3HMYstY6ljffqanFgY9SkLfL/WH37iqHywcKBnomhYUJkdTSN7A5+n4vAPFYtNkV0po4IIMh+Fp03/gaqUu8RIyq6okNrCI5L3W7BaMlhvKmOpkphKt1LJ4na37th71tm1fdbk1YLgz1TgZKOPUgTpdck1+bNM8B4dv5M4Rk/wTod2ZheJXXN8LZa6yLR7FQBCM19ZLU+8FPfT67r/hDwAA//8DAFBLAQItABQABgAIAAAAIQCREKFLmQEAAPgGAAATAAAAAAAAAAAAAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhABNevmUCAQAA3wIAAAsAAAAAAAAAAAAAAAAA0gMAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAGl5S20UBAAAsgkAAA8AAAAAAAAAAAAAAAAABQcAAHhsL3dvcmtib29rLnhtbFBLAQItABQABgAIAAAAIQAuZhB6PAEAAOEEAAAaAAAAAAAAAAAAAAAAAEYLAAB4bC9fcmVscy93b3JrYm9vay54bWwucmVsc1BLAQItABQABgAIAAAAIQBMgXCbvAkAAMZIAAAYAAAAAAAAAAAAAAAAAMINAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWxQSwECLQAUAAYACAAAACEAkwlHQMEHAAATIgAAEwAAAAAAAAAAAAAAAAC0FwAAeGwvdGhlbWUvdGhlbWUxLnhtbFBLAQItABQABgAIAAAAIQBKjEHqKgMAACoIAAANAAAAAAAAAAAAAAAAAKYfAAB4bC9zdHlsZXMueG1sUEsBAi0AFAAGAAgAAAAhACZVd1jtAwAA+iIAABQAAAAAAAAAAAAAAAAA+yIAAHhsL3NoYXJlZFN0cmluZ3MueG1sUEsBAi0AFAAGAAgAAAAhAK0x1F2kAAAA2gAAABUAAAAAAAAAAAAAAAAAGicAAHhsL3BlcnNvbnMvcGVyc29uLnhtbFBLAQItABQABgAIAAAAIQCQMJoDPwEAAEsCAAATAAAAAAAAAAAAAAAAAPEnAABjdXN0b21YbWwvaXRlbTEueG1sUEsBAi0AFAAGAAgAAAAhAAUiBzw+AQAAIwIAABgAAAAAAAAAAAAAAAAAiSkAAGN1c3RvbVhtbC9pdGVtUHJvcHMxLnhtbFBLAQItABQABgAIAAAAIQC9hGIjkAAAANsAAAATAAAAAAAAAAAAAAAAACUrAABjdXN0b21YbWwvaXRlbTIueG1sUEsBAi0AFAAGAAgAAAAhAD0oSnXwAAAATwEAABgAAAAAAAAAAAAAAAAADiwAAGN1c3RvbVhtbC9pdGVtUHJvcHMyLnhtbFBLAQItABQABgAIAAAAIQCst+yL4AoAAPg5AAATAAAAAAAAAAAAAAAAAFwtAABjdXN0b21YbWwvaXRlbTMueG1sUEsBAi0AFAAGAAgAAAAhAJqUUKW7AQAAfQQAABgAAAAAAAAAAAAAAAAAlTgAAGN1c3RvbVhtbC9pdGVtUHJvcHMzLnhtbFBLAQItABQABgAIAAAAIQA4SkXsYwEAAH4CAAARAAAAAAAAAAAAAAAAAK46AABkb2NQcm9wcy9jb3JlLnhtbFBLAQItABQABgAIAAAAIQAnAeylkQEAABcDAAAQAAAAAAAAAAAAAAAAAEg9AABkb2NQcm9wcy9hcHAueG1sUEsBAi0AFAAGAAgAAAAhAFlJ3sosAQAAEgIAABMAAAAAAAAAAAAAAAAAD0AAAGRvY1Byb3BzL2N1c3RvbS54bWxQSwECLQAUAAYACAAAACEAdD85esIAAAAoAQAAHgAAAAAAAAAAAAAAAAB0QgAAY3VzdG9tWG1sL19yZWxzL2l0ZW0xLnhtbC5yZWxzUEsBAi0AFAAGAAgAAAAhAFyWJyLDAAAAKAEAAB4AAAAAAAAAAAAAAAAAekQAAGN1c3RvbVhtbC9fcmVscy9pdGVtMi54bWwucmVsc1BLAQItABQABgAIAAAAIQB78wKjwwAAACgBAAAeAAAAAAAAAAAAAAAAAIFGAABjdXN0b21YbWwvX3JlbHMvaXRlbTMueG1sLnJlbHNQSwUGAAAAABUAFQB9BQAAiEgAAAAA"""

@st.cache_data
def cargar_mapeo_monedas():
    """Carga el mapeo de monedas desde el archivo Excel embebido actualizado"""
    try:
        # Decodificar el archivo base64
        archivo_bytes = base64.b64decode(ARCHIVO_MAPEO_MONEDAS_BASE64)
        
        # Leer el Excel desde memoria
        df_mapeo = pd.read_excel(io.BytesIO(archivo_bytes), engine='openpyxl')
        
        # Crear diccionario de mapeo
        mapeo_dict = dict(zip(df_mapeo['Flex Efectivo'].astype(str), df_mapeo['Moneda'].astype(str)))
        
        # Log para debug
        print(f"✅ Mapeo de monedas cargado: {len(mapeo_dict)} entradas")
        print(f"✅ Monedas disponibles: {sorted(df_mapeo['Moneda'].unique())}")
        
        return mapeo_dict, df_mapeo
        
    except Exception as e:
        st.warning(f"Error cargando mapeo actualizado: {e}")
        # Mapeo de respaldo básico
        mapeo_respaldo = {
            'USD': 'USD', 'EUR': 'EUR', 'PEN': 'PEN', 'CLP': 'CLP', 
            'BRL': 'BRL', 'MXN': 'MXN', 'ARS': 'ARS', 'COP': 'COP', 'UYU': 'UYU'
        }
        return mapeo_respaldo, None

def limpiar_nombre_banco(nombre_hoja):
    """Extrae el nombre del banco desde el nombre de la hoja"""
    nombre_limpio = re.sub(r'\d+$', '', nombre_hoja)
    nombre_limpio = re.sub(r'\b(datos|data|info|información)\b', '', nombre_limpio, flags=re.IGNORECASE)
    nombre_limpio = nombre_limpio.strip().strip('_-. ')
    return nombre_limpio if nombre_limpio else nombre_hoja

def obtener_moneda_de_flex(flex_banco, mapeo_monedas):
    """
    Obtiene la moneda basada en el flex banco usando mapeo actualizado
    Implementa búsqueda exacta y por prefijos
    """
    if pd.isna(flex_banco) or flex_banco == '':
        return ''
    
    flex_str = str(flex_banco).strip()
    
    # 1. Búsqueda exacta
    if flex_str in mapeo_monedas:
        return mapeo_monedas[flex_str]
    
    # 2. Búsqueda por coincidencia parcial (contiene)
    for flex_clave, moneda in mapeo_monedas.items():
        if flex_str in flex_clave or flex_clave in flex_str:
            return moneda
    
    # 3. Búsqueda por prefijos (para claves jerárquicas como 113.11121.1303...)
    partes_flex = flex_str.split('.')
    for i in range(len(partes_flex), 0, -1):
        prefijo = '.'.join(partes_flex[:i])
        for flex_clave, moneda in mapeo_monedas.items():
            if flex_clave.startswith(prefijo):
                return moneda
    
    # 4. Búsqueda por código de moneda contenido (fallback)
    flex_upper = flex_str.upper()
    codigos_moneda = ['USD', 'EUR', 'PEN', 'CLP', 'BRL', 'MXN', 'ARS', 'COP', 'UYU']
    for codigo in codigos_moneda:
        if codigo in flex_upper:
            return codigo
    
    return ''

def normalizar_fecha_sin_hora(fecha_value):
    """
    Normaliza diferentes tipos de fecha para eliminar la hora y dejar solo la fecha
    """
    try:
        if pd.isna(fecha_value) or fecha_value is None:
            return ''
        
        if hasattr(fecha_value, 'date') and callable(fecha_value.date):
            return fecha_value.date().strftime('%Y-%m-%d')
        
        if hasattr(fecha_value, 'strftime'):
            return fecha_value.strftime('%Y-%m-%d')
        
        fecha_str = str(fecha_value).strip()
        
        if ' ' in fecha_str and fecha_str != 'nan':
            try:
                fecha_parseada = pd.to_datetime(fecha_str, errors='coerce')
                if not pd.isna(fecha_parseada):
                    return fecha_parseada.strftime('%Y-%m-%d')
            except:
                pass
        
        if fecha_str and fecha_str != 'nan':
            try:
                fecha_parseada = pd.to_datetime(fecha_str, errors='coerce')
                if not pd.isna(fecha_parseada):
                    return fecha_parseada.strftime('%Y-%m-%d')
            except:
                pass
        
        return fecha_str if fecha_str != 'nan' else ''
    
    except Exception:
        return ''

def procesar_archivo_excel(archivo, progress_callback=None, log_callback=None):
    """Procesa el archivo Excel y devuelve los datos consolidados"""
    
    def log(mensaje):
        if log_callback:
            log_callback(mensaje)
        print(mensaje)
    
    def actualizar_progreso(valor):
        if progress_callback:
            progress_callback(valor)
    
    try:
        log("🔍 Iniciando procesamiento del archivo...")
        
        # Cargar el mapeo de monedas actualizado
        mapeo_monedas, df_mapeo_original = cargar_mapeo_monedas()
        log(f"✅ Mapeo de monedas actualizado cargado: {len(mapeo_monedas)} entradas")
        if df_mapeo_original is not None:
            monedas_disponibles = sorted(df_mapeo_original['Moneda'].unique())
            log(f"💰 Monedas disponibles: {monedas_disponibles}")
        
        # Leer todas las hojas del archivo
        try:
            todas_las_hojas = pd.read_excel(archivo, sheet_name=None, header=None, engine='openpyxl')
            log(f"📄 Archivo cargado: {len(todas_las_hojas)} hojas encontradas")
        except Exception as e:
            log(f"❌ Error al leer el archivo: {e}")
            return None, None, None, None
        
        # Filtrar hojas que no son de datos
        hojas_excluidas = ['resumen', 'summary', 'índice', 'index', 'instrucciones', 'instructions', 'totales', 'total', 'config', 'configuración']
        hojas_datos = []
        
        for nombre_hoja in todas_las_hojas.keys():
            if not any(excl.lower() in nombre_hoja.lower() for excl in hojas_excluidas):
                hojas_datos.append(nombre_hoja)
        
        log(f"📊 Hojas de datos detectadas: {hojas_datos}")
        actualizar_progreso(10)
        
        # Procesar cada hoja
        datos_consolidados = []
        resumen_proceso = []
        total_hojas = len(hojas_datos)
        
        for i, nombre_hoja in enumerate(hojas_datos):
            try:
                log(f"\n🔄 Procesando hoja '{nombre_hoja}'...")
                
                df_hoja = todas_las_hojas[nombre_hoja]
                
                if len(df_hoja) < 13:
                    log(f"⚠️  Hoja '{nombre_hoja}' tiene pocas filas ({len(df_hoja)}), omitiendo...")
                    resumen_proceso.append({
                        'Banco': limpiar_nombre_banco(nombre_hoja),
                        'Hoja': nombre_hoja,
                        'Registros': 0,
                        'Estado': 'Omitida - Pocas filas'
                    })
                    continue
                
                # Extraer encabezados desde la fila 11 (índice 10)
                encabezados = df_hoja.iloc[10].fillna('').astype(str)
                
                for j, enc in enumerate(encabezados):
                    if enc == '' or enc == 'nan':
                        encabezados.iloc[j] = f'Col_{j}'
                
                # Mapear columnas por nombre
                mapeo_columnas = {}
                for idx, nombre in enumerate(encabezados):
                    nombre_lower = str(nombre).lower().strip()
                    
                    if 'estado' in nombre_lower:
                        mapeo_columnas['Estado'] = idx
                    elif 'aging' in nombre_lower:
                        mapeo_columnas['Aging'] = idx
                    elif 'fecha' in nombre_lower and 'contable' in nombre_lower:
                        mapeo_columnas['Fecha'] = idx
                    elif 'fecha' in nombre_lower and 'transac' in nombre_lower:
                        mapeo_columnas['Fecha transacción'] = idx
                    elif 'categoría' in nombre_lower or 'categoria' in nombre_lower:
                        mapeo_columnas['Categoría'] = idx
                    elif 'monto' in nombre_lower and ('funcional' in nombre_lower or 'total' in nombre_lower):
                        mapeo_columnas['Monto'] = idx
                    elif 'concepto' in nombre_lower:
                        mapeo_columnas['Concepto'] = idx
                    elif 'responsable' in nombre_lower:
                        mapeo_columnas['Responsable'] = idx
                    elif 'flex' in nombre_lower and 'contable' in nombre_lower:
                        mapeo_columnas['Flex contable'] = idx
                    elif 'flex' in nombre_lower and ('banco' in nombre_lower or 'efectivo' in nombre_lower):
                        mapeo_columnas['Flex banco'] = idx
                    elif 'tipo' in nombre_lower and 'extracto' in nombre_lower:
                        mapeo_columnas['Tipo extracto'] = idx
                
                # Mapeos fijos por índice
                if len(encabezados) > 7:
                    mapeo_columnas['Numero de transacción'] = 7
                if len(encabezados) > 11:
                    mapeo_columnas['Proveedor/Cliente'] = 11
                
                log(f"📋 Columnas mapeadas: {list(mapeo_columnas.keys())}")
                
                # Procesar filas desde el índice 12
                filas_validas = []
                
                for fila_idx in range(12, len(df_hoja)):
                    fila = df_hoja.iloc[fila_idx]
                    criterios_cumplidos = 0
                    
                    # Criterio 1: Aging numérico > 0
                    if 'Aging' in mapeo_columnas:
                        try:
                            aging_val = pd.to_numeric(fila.iloc[mapeo_columnas['Aging']], errors='coerce')
                            if not pd.isna(aging_val) and aging_val > 0:
                                criterios_cumplidos += 1
                        except:
                            pass
                    
                    # Criterio 2: Fecha no vacía
                    fecha_encontrada = False
                    for col_fecha in ['Fecha', 'Fecha transacción']:
                        if col_fecha in mapeo_columnas:
                            fecha_val = str(fila.iloc[mapeo_columnas[col_fecha]])
                            if fecha_val and fecha_val != 'nan' and fecha_val.strip():
                                fecha_encontrada = True
                                break
                    if fecha_encontrada:
                        criterios_cumplidos += 1
                    
                    # Criterio 3: Monto numérico distinto de 0
                    if 'Monto' in mapeo_columnas:
                        try:
                            monto_val = pd.to_numeric(fila.iloc[mapeo_columnas['Monto']], errors='coerce')
                            if not pd.isna(monto_val) and monto_val != 0:
                                criterios_cumplidos += 1
                        except:
                            pass
                    
                    # Criterio 4: Responsable con longitud > 2
                    if 'Responsable' in mapeo_columnas:
                        responsable_val = str(fila.iloc[mapeo_columnas['Responsable']])
                        if responsable_val and responsable_val != 'nan' and len(responsable_val.strip()) > 2:
                            criterios_cumplidos += 1
                    
                    # Criterio 5: Flex contable que contenga '105.' y longitud > 10
                    if 'Flex contable' in mapeo_columnas:
                        flex_cont_val = str(fila.iloc[mapeo_columnas['Flex contable']])
                        if flex_cont_val and flex_cont_val != 'nan' and '105.' in flex_cont_val and len(flex_cont_val.strip()) > 10:
                            criterios_cumplidos += 1
                    
                    if criterios_cumplidos >= 3:
                        filas_validas.append(fila_idx)
                
                if not filas_validas:
                    log(f"⚠️  No se encontraron filas válidas en '{nombre_hoja}'")
                    resumen_proceso.append({
                        'Banco': limpiar_nombre_banco(nombre_hoja),
                        'Hoja': nombre_hoja,
                        'Registros': 0,
                        'Estado': 'Sin datos válidos'
                    })
                    continue
                
                # Crear DataFrame con las filas válidas
                df_procesado = pd.DataFrame()
                
                columnas_objetivo = [
                    'Estado', 'Aging', 'Fecha', 'Fecha transacción', 'Categoría', 
                    'Numero de transacción', 'Proveedor/Cliente', 'Monto', 'Concepto', 
                    'Responsable', 'Flex contable', 'Flex banco', 'Tipo extracto',
                    'Moneda', 'BANCO'
                ]
                
                for col in columnas_objetivo[:-2]:
                    if col in mapeo_columnas:
                        valores = [df_hoja.iloc[idx, mapeo_columnas[col]] if idx < len(df_hoja) and mapeo_columnas[col] < len(df_hoja.columns) else '' for idx in filas_validas]
                        df_procesado[col] = valores
                    else:
                        df_procesado[col] = [''] * len(filas_validas)
                
                # Agregar columna de moneda usando mapeo actualizado
                monedas_asignadas = 0
                if 'Flex banco' in mapeo_columnas:
                    monedas_aplicadas = []
                    for flex_val in df_procesado['Flex banco']:
                        moneda = obtener_moneda_de_flex(flex_val, mapeo_monedas)
                        monedas_aplicadas.append(moneda)
                        if moneda:
                            monedas_asignadas += 1
                    
                    df_procesado['Moneda'] = monedas_aplicadas
                    log(f"💰 Monedas asignadas: {monedas_asignadas} de {len(filas_validas)} registros")
                else:
                    df_procesado['Moneda'] = ''
                
                # Agregar columna de banco
                df_procesado['BANCO'] = limpiar_nombre_banco(nombre_hoja)
                
                # Limpiar datos
                for col in df_procesado.columns:
                    if df_procesado[col].dtype == 'object':
                        df_procesado[col] = df_procesado[col].astype(str).str.strip()
                        df_procesado[col] = df_procesado[col].replace('nan', '')
                
                datos_consolidados.append(df_procesado)
                
                log(f"✅ Hoja '{nombre_hoja}' procesada: {len(filas_validas)} registros válidos")
                resumen_proceso.append({
                    'Banco': limpiar_nombre_banco(nombre_hoja),
                    'Hoja': nombre_hoja,
                    'Registros': len(filas_validas),
                    'Estado': 'Procesada correctamente'
                })
                
                progreso = 10 + (60 * (i + 1) / total_hojas)
                actualizar_progreso(int(progreso))
                
            except Exception as e:
                log(f"❌ Error procesando hoja '{nombre_hoja}': {e}")
                resumen_proceso.append({
                    'Banco': limpiar_nombre_banco(nombre_hoja),
                    'Hoja': nombre_hoja,
                    'Registros': 0,
                    'Estado': f'Error: {str(e)[:50]}'
                })
        
        if not datos_consolidados:
            log("❌ No se pudo procesar ninguna hoja")
            return None, None, None, None
        
        # Consolidar todos los datos
        log("\n🔗 Consolidando datos...")
        df_final = pd.concat(datos_consolidados, ignore_index=True)
        actualizar_progreso(75)
        
        # Limpiar datos finales
        for col in df_final.columns:
            if df_final[col].dtype == 'object':
                df_final[col] = df_final[col].astype(str).str.strip()
                df_final[col] = df_final[col].replace('nan', '')
                if col not in ['Proveedor/Cliente', 'Moneda']:
                    df_final[col] = df_final[col].replace('', np.nan)
        
        # Normalizar fechas sin hora
        log("📅 Normalizando fechas sin hora...")
        if 'Fecha' in df_final.columns:
            fechas_procesadas = 0
            for idx in df_final.index:
                fecha_original = df_final.at[idx, 'Fecha']
                fecha_normalizada = normalizar_fecha_sin_hora(fecha_original)
                df_final.at[idx, 'Fecha'] = fecha_normalizada
                if fecha_normalizada:
                    fechas_procesadas += 1
            
            log(f"✅ Fechas normalizadas: {fechas_procesadas} registros procesados")
        
        actualizar_progreso(85)
        
        # Calcular DEBE/HABER/SALDO
        log("💰 Calculando DEBE, HABER y SALDO...")
        
        col_estado = None
        col_monto = None
        col_tipo_extracto = None
        
        for col in df_final.columns:
            if 'estado' in col.lower():
                col_estado = col
            if 'monto' in col.lower():
                col_monto = col
            if 'tipo' in col.lower() and 'extracto' in col.lower():
                col_tipo_extracto = col
        
        log(f"🔍 Columnas identificadas - Estado: {col_estado}, Monto: {col_monto}, Tipo extracto: {col_tipo_extracto}")
        
        if col_estado and col_monto:
            df_final['DEBE'] = 0.0
            df_final['HABER'] = 0.0
            
            textos_debe = [
                "III. Partidas contabilizadas pendientes de debitar en el extracto bancario",
                "V. Partidas acreditadas en el extracto bancario, pendientes de contabilizar"
            ]
            
            textos_haber = [
                "II. Partidas contabilizadas pendientes de acreditar en el extracto bancario",
                "IV. Partidas debitadas en el extracto bancario, pendientes de contabilizar"
            ]
            
            debe_count = 0
            haber_count = 0
            
            for idx, row in df_final.iterrows():
                estado_text = str(row[col_estado]).strip()
                tipo_extracto = str(row.get(col_tipo_extracto, '')).strip().upper() if col_tipo_extracto else ''
                
                try:
                    monto_val = pd.to_numeric(row[col_monto], errors='coerce')
                    if pd.isna(monto_val):
                        monto_val = 0
                except:
                    monto_val = 0
                
                if tipo_extracto == 'DEBIT':
                    df_final.at[idx, 'DEBE'] = abs(monto_val)
                    df_final.at[idx, 'HABER'] = 0.0
                    debe_count += 1
                elif tipo_extracto == 'CREDIT':
                    df_final.at[idx, 'HABER'] = abs(monto_val)
                    df_final.at[idx, 'DEBE'] = 0.0
                    haber_count += 1
                elif any(texto in estado_text for texto in textos_debe):
                    df_final.at[idx, 'DEBE'] = abs(monto_val)
                    df_final.at[idx, 'HABER'] = 0.0
                    debe_count += 1
                elif any(texto in estado_text for texto in textos_haber):
                    df_final.at[idx, 'HABER'] = abs(monto_val)
                    df_final.at[idx, 'DEBE'] = 0.0
                    haber_count += 1
                else:
                    if monto_val > 0:
                        df_final.at[idx, 'DEBE'] = monto_val
                        df_final.at[idx, 'HABER'] = 0.0
                        debe_count += 1
                    elif monto_val < 0:
                        df_final.at[idx, 'HABER'] = abs(monto_val)
                        df_final.at[idx, 'DEBE'] = 0.0
                        haber_count += 1
            
            df_final['SALDO'] = df_final['DEBE'] - df_final['HABER']
            log(f"💰 Cálculo completado: {debe_count} registros en DEBE, {haber_count} en HABER")
            
        else:
            log("⚠️  No se pudieron identificar columnas de Estado o Monto para calcular DEBE/HABER")
            df_final['DEBE'] = 0.0
            df_final['HABER'] = 0.0
            df_final['SALDO'] = 0.0
        
        # Reordenar columnas
        columnas_finales = [
            'Estado', 'Aging', 'Fecha', 'Fecha transacción', 'Categoría', 
            'Numero de transacción', 'Proveedor/Cliente', 'Monto', 'Concepto', 
            'Responsable', 'Flex contable', 'Flex banco', 'Tipo extracto',
            'Moneda', 'BANCO', 'DEBE', 'HABER', 'SALDO'
        ]
        
        for col in df_final.columns:
            if col not in columnas_finales:
                columnas_finales.append(col)
        
        df_final = df_final[[col for col in columnas_finales if col in df_final.columns]]
        
        # Crear DataFrames de resumen
        df_resumen = pd.DataFrame(resumen_proceso)
        
        # Estadísticas de monedas
        if 'Moneda' in df_final.columns:
            monedas_stats = df_final['Moneda'].value_counts()
            df_monedas = pd.DataFrame({
                'Moneda': monedas_stats.index,
                'Registros': monedas_stats.values
            }).reset_index(drop=True)
        else:
            df_monedas = pd.DataFrame()
        
        df_estadisticas = pd.DataFrame({
            'Métrica': [
                'Total de registros procesados',
                'Número de bancos procesados', 
                'Total DEBE',
                'Total HABER',
                'SALDO FINAL',
                'Monedas diferentes encontradas',
                'Fecha de procesamiento',
                'Archivo procesado'
            ],
            'Valor': [
                len(df_final),
                len(df_final['BANCO'].unique()) if 'BANCO' in df_final.columns else 0,
                f"{df_final['DEBE'].sum():,.2f}" if 'DEBE' in df_final.columns else 0,
                f"{df_final['HABER'].sum():,.2f}" if 'HABER' in df_final.columns else 0,
                f"{df_final['SALDO'].sum():,.2f}" if 'SALDO' in df_final.columns else 0,
                df_final['Moneda'].nunique() if 'Moneda' in df_final.columns else 0,
                datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
                'Archivo subido'
            ]
        })
        
        actualizar_progreso(100)
        log(f"\n🎉 ¡Procesamiento completado exitosamente!")
        log(f"📊 Total de registros: {len(df_final)}")
        log(f"🏦 Bancos procesados: {len(df_final['BANCO'].unique()) if 'BANCO' in df_final.columns else 0}")
        log(f"💰 Total DEBE: {df_final['DEBE'].sum():,.2f}")
        log(f"💰 Total HABER: {df_final['HABER'].sum():,.2f}")
        log(f"💰 SALDO: {df_final['SALDO'].sum():,.2f}")
        log(f"💱 Monedas procesadas: {df_final['Moneda'].nunique() if 'Moneda' in df_final.columns else 0}")
        
        return df_final, df_resumen, df_estadisticas, df_monedas
        
    except Exception as e:
        log(f"❌ Error crítico durante el procesamiento: {e}")
        return None, None, None, None

def main():
    """Función principal de la aplicación Streamlit"""
    
    # Título y descripción
    st.title("🏦 Vaciado de Carátulas Bancarias")
    st.markdown("---")
    st.markdown("""
    ### 📋 Instrucciones de uso:
    1. **Sube tu archivo Excel** con las carátulas bancarias
    2. **Haz clic en 'Procesar archivo'** para iniciar el análisis
    3. **Descarga los resultados** en formato Excel
    
    ⚠️  **Importante**: El archivo debe tener encabezados en la fila 11 y datos a partir de la fila 13.
    
    ✨ **Actualizaciones**:
    - 📅 Fechas sin hora (formato YYYY-MM-DD)
    - 💱 **Mapeo de monedas actualizado** con 136 entradas
    - 🎯 Monedas soportadas: USD, EUR, PEN, CLP, BRL, MXN, ARS, COP, UYU
    """)
    
    # Sidebar con información del mapeo
    st.sidebar.header("⚙️ Configuración")
    st.sidebar.info("""
    **Formato esperado:**
    - Encabezados en fila 11
    - Datos desde fila 13
    - Columnas: Estado, Aging, Fecha, Monto, etc.
    """)
    
    # Mostrar información del mapeo de monedas
    with st.sidebar.expander("💱 Mapeo de Monedas"):
        try:
            mapeo_monedas, df_mapeo = cargar_mapeo_monedas()
            if df_mapeo is not None:
                st.write(f"**Total de entradas:** {len(mapeo_monedas)}")
                st.write(f"**Monedas disponibles:** {len(df_mapeo['Moneda'].unique())}")
                monedas_lista = sorted(df_mapeo['Moneda'].unique())
                st.write("**Códigos:**")
                for moneda in monedas_lista:
                    st.write(f"• {moneda}")
            else:
                st.write("Mapeo básico de respaldo cargado")
        except Exception as e:
            st.write(f"Error: {e}")
    
    # Upload de archivo
    archivo_subido = st.file_uploader(
        "📁 Selecciona el archivo Excel con las carátulas",
        type=['xlsx', 'xlsm', 'xls'],
        help="Formatos soportados: .xlsx, .xlsm, .xls"
    )
    
    if archivo_subido is not None:
        st.success(f"✅ Archivo cargado: {archivo_subido.name}")
        
        # Mostrar información del archivo
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📄 Nombre", archivo_subido.name)
        with col2:
            st.metric("📦 Tamaño", f"{archivo_subido.size / 1024:.1f} KB")
        with col3:
            st.metric("📋 Tipo", archivo_subido.type)
        
        # Botón para procesar
        if st.button("🚀 Procesar archivo", type="primary"):
            
            # Contenedores para progreso y logs
            progress_container = st.container()
            log_container = st.container()
            
            with progress_container:
                st.subheader("📊 Progreso del procesamiento")
                progress_bar = st.progress(0)
                status_text = st.empty()
            
            with log_container:
                st.subheader("📝 Registro de actividad")
                log_area = st.empty()
            
            # Variables para logs
            logs = []
            
            def actualizar_progreso(valor):
                progress_bar.progress(valor)
                status_text.text(f"Progreso: {valor}%")
            
            def agregar_log(mensaje):
                logs.append(mensaje)
                log_area.text_area(
                    "Logs:", 
                    value="\n".join(logs), 
                    height=200,
                    key=f"logs_{len(logs)}"
                )
            
            # Procesar archivo
            try:
                df_final, df_resumen, df_estadisticas, df_monedas = procesar_archivo_excel(
                    archivo_subido,
                    progress_callback=actualizar_progreso,
                    log_callback=agregar_log
                )
                
                if df_final is not None:
                    st.success("🎉 ¡Archivo procesado exitosamente!")
                    
                    # Mostrar estadísticas
                    st.subheader("📈 Estadísticas del procesamiento")
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.metric("📊 Registros", len(df_final))
                    with col2:
                        bancos_unicos = len(df_final['BANCO'].unique()) if 'BANCO' in df_final.columns else 0
                        st.metric("🏦 Bancos", bancos_unicos)
                    with col3:
                        suma_debe = df_final['DEBE'].sum() if 'DEBE' in df_final.columns else 0
                        st.metric("💰 Total DEBE", f"{suma_debe:,.2f}")
                    with col4:
                        suma_haber = df_final['HABER'].sum() if 'HABER' in df_final.columns else 0
                        st.metric("💰 Total HABER", f"{suma_haber:,.2f}")
                    
                    # Mostrar SALDO final y monedas
                    col1, col2 = st.columns(2)
                    with col1:
                        saldo_final = suma_debe - suma_haber
                        if saldo_final > 0:
                            st.success(f"💰 **SALDO FINAL: ${saldo_final:,.2f}** (A favor)")
                        elif saldo_final < 0:
                            st.error(f"💰 **SALDO FINAL: ${abs(saldo_final):,.2f}** (En contra)")
                        else:
                            st.info(f"💰 **SALDO FINAL: $0.00** (Balanceado)")
                    
                    with col2:
                        if 'Moneda' in df_final.columns:
                            monedas_procesadas = df_final['Moneda'].nunique()
                            st.info(f"💱 **{monedas_procesadas} monedas diferentes** procesadas")
                    
                    # Tabs para mostrar resultados
                    tab1, tab2, tab3, tab4, tab5 = st.tabs([
                        "📋 Datos Consolidados", 
                        "💱 Análisis de Monedas", 
                        "📅 Vista de Fechas", 
                        "📊 Resumen del Proceso", 
                        "📈 Estadísticas"
                    ])
                    
                    with tab1:
                        st.subheader("📋 Datos Consolidados")
                        st.dataframe(df_final, use_container_width=True)
                        st.info(f"Total de registros: {len(df_final)}")
                        
                        # Mostrar muestra por DEBE/HABER
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            debe_registros = df_final[df_final['DEBE'] > 0]
                            if len(debe_registros) > 0:
                                st.subheader(f"💰 DEBE ({len(debe_registros)} registros)")
                                st.dataframe(
                                    debe_registros[['Estado', 'Fecha', 'Moneda', 'DEBE', 'BANCO']].head(5), 
                                    use_container_width=True
                                )
                        
                        with col2:
                            haber_registros = df_final[df_final['HABER'] > 0]
                            if len(haber_registros) > 0:
                                st.subheader(f"💰 HABER ({len(haber_registros)} registros)")
                                st.dataframe(
                                    haber_registros[['Estado', 'Fecha', 'Moneda', 'HABER', 'BANCO']].head(5), 
                                    use_container_width=True
                                )
                    
                    with tab2:
                        st.subheader("💱 Análisis de Monedas")
                        
                        if df_monedas is not None and len(df_monedas) > 0:
                            st.dataframe(df_monedas, use_container_width=True)
                            
                            # Gráfico de monedas
                            st.subheader("📊 Distribución por Moneda")
                            st.bar_chart(df_monedas.set_index('Moneda'))
                            
                            # Detalle por moneda
                            st.subheader("💰 Totales por Moneda")
                            if 'Moneda' in df_final.columns:
                                moneda_totales = df_final.groupby('Moneda').agg({
                                    'DEBE': 'sum',
                                    'HABER': 'sum',
                                    'SALDO': 'sum'
                                }).round(2)
                                st.dataframe(moneda_totales, use_container_width=True)
                        else:
                            st.warning("No se encontraron monedas asignadas")
                            
                            # Mostrar registros sin moneda
                            sin_moneda = df_final[df_final['Moneda'] == '']
                            if len(sin_moneda) > 0:
                                st.subheader("⚠️ Registros sin moneda asignada")
                                st.dataframe(sin_moneda[['Flex banco', 'BANCO', 'Monto']].head(10), use_container_width=True)
                    
                    with tab3:
                        st.subheader("📅 Vista de Fechas Normalizadas")
                        if 'Fecha' in df_final.columns:
                            fechas_df = df_final[['Fecha', 'BANCO', 'Moneda', 'Monto']].copy()
                            fechas_df = fechas_df[fechas_df['Fecha'].notna() & (fechas_df['Fecha'] != '')]
                            
                            if len(fechas_df) > 0:
                                st.dataframe(fechas_df.head(20), use_container_width=True)
                                
                                col1, col2, col3 = st.columns(3)
                                with col1:
                                    st.metric("📅 Total con fecha", len(fechas_df))
                                with col2:
                                    st.metric("📅 Fechas únicas", fechas_df['Fecha'].nunique())
                                with col3:
                                    if len(fechas_df) > 0:
                                        st.metric("📅 Formato", fechas_df['Fecha'].iloc[0])
                            else:
                                st.warning("No se encontraron fechas válidas")
                        else:
                            st.warning("No se encontró columna de Fecha")
                    
                    with tab4:
                        st.subheader("📊 Resumen del Proceso")
                        if df_resumen is not None:
                            st.dataframe(df_resumen, use_container_width=True)
                        else:
                            st.warning("No hay datos de resumen disponibles")
                    
                    with tab5:
                        st.subheader("📈 Estadísticas Generales")
                        if df_estadisticas is not None:
                            st.dataframe(df_estadisticas, use_container_width=True)
                        else:
                            st.warning("No hay estadísticas disponibles")
                    
                    # Generar archivo Excel para descarga
                    st.subheader("💾 Descargar resultados")
                    
                    # Crear archivo Excel en memoria
                    output = io.BytesIO()
                    
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_final.to_excel(writer, sheet_name='Datos_Consolidados', index=False)
                        if df_resumen is not None:
                            df_resumen.to_excel(writer, sheet_name='Resumen_Proceso', index=False)
                        if df_estadisticas is not None:
                            df_estadisticas.to_excel(writer, sheet_name='Estadisticas', index=False)
                        if df_monedas is not None and len(df_monedas) > 0:
                            df_monedas.to_excel(writer, sheet_name='Analisis_Monedas', index=False)
                    
                    # Preparar archivo para descarga
                    excel_data = output.getvalue()
                    fecha_actual = datetime.now().strftime('%Y%m%d_%H%M%S')
                    nombre_archivo = f'caratulas_vaciado_{fecha_actual}.xlsx'
                    
                    st.download_button(
                        label="📥 Descargar archivo Excel",
                        data=excel_data,
                        file_name=nombre_archivo,
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        type="primary"
                    )
                    
                    st.success("✅ ¡Listo! Haz clic en el botón de arriba para descargar el archivo procesado.")
                    
                else:
                    st.error("❌ No se pudo procesar el archivo. Revisa los logs para más información.")
                    
            except Exception as e:
                st.error(f"❌ Error durante el procesamiento: {e}")
                agregar_log(f"Error crítico: {e}")
    
    # Footer
    st.markdown("---")
    st.markdown("🔧 **Desarrollado para automatizar el procesamiento de carátulas bancarias**")
    st.markdown("💱 **Mapeo de monedas actualizado**: 136 entradas de Flex → Moneda")
    st.markdown("📅 **Fechas normalizadas**: Sin hora (formato YYYY-MM-DD)")
    st.markdown("🎯 **Monedas soportadas**: USD, EUR, PEN, CLP, BRL, MXN, ARS, COP, UYU")

if __name__ == "__main__":
    main()
