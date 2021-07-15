using SBO.Hub.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http.Headers;

namespace SBO.Hub.Util
{
    public class APICallUtil
    {
        public string ApiUrl;
        public string MediaType;

        public APICallUtil(string apiUrl = "", string mediaType = "application/json")
        {
            MediaType = mediaType;
            if (String.IsNullOrEmpty(apiUrl))
            {
                ApiUrl = System.Configuration.ConfigurationManager.AppSettings["WebApiURL"];
            }
            else
            {
                ApiUrl = apiUrl;
            }
        }

        public T Get<T>(string url) where T : class
        {
            return Task.Run(async () =>
            {
                return await GetAsync<T>(url);
            }).Result;
        }

        public string Post<T>(string controller, T item) where T : class
        {
            return Task.Run(async () =>
            {
                return await PostAsync(controller, item);
            }).Result;
        }

        public async Task<T> GetAsync<T>(string url) where T : class
        {
            string JsonResult = "";
            try
            {
                var client = new HttpClient();

                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(MediaType));
                var response = await client.GetAsync(url);

                JsonResult = response.Content.ReadAsStringAsync().Result;
                var rootobject = JsonConvert.DeserializeObject<T>(JsonResult);

                return rootobject;
            }
            catch (Exception e)
            {
                try
                {
                    if (!String.IsNullOrEmpty(JsonResult))
                    {
                        APIErrorModel errorModel = JsonConvert.DeserializeObject<APIErrorModel>(JsonResult);
                        throw new Exception(errorModel.Message, new Exception(errorModel.ExceptionMessage));
                    }
                    else throw e;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        /// <summary>
        /// Retorna o model de acordo com o parâmetro passado
        /// </summary>
        /// <typeparam name="T">Tipo do model</typeparam>
        /// <param name="controller">Nome do controller da web api</param>
        /// <param name="param">Parâmetro(s) de busca do model</param>
        /// <returns></returns>
        public async Task<T> GetAsync<T>(string controller, string param, string paramSeparator = "/") where T : class
        {
            string JsonResult = "";
            try
            {
                string url = string.Format(ApiUrl + "/{0}{1}{2}", controller, paramSeparator, param);

                var client = new HttpClient();

                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(MediaType));
                var response = await client.GetAsync(url);

                JsonResult = response.Content.ReadAsStringAsync().Result;
                var rootobject = JsonConvert.DeserializeObject<T>(JsonResult);

                return rootobject;
            }
            catch (Exception e)
            {
                try
                {
                    if (!String.IsNullOrEmpty(JsonResult))
                    {
                        APIErrorModel errorModel = JsonConvert.DeserializeObject<APIErrorModel>(JsonResult);
                        throw new Exception(errorModel.Message, new Exception(errorModel.ExceptionMessage));
                    }
                    else throw e;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        /// <summary>
        /// Retorna uma lista de model de acordo com o parâmetro passado
        /// </summary>
        /// <typeparam name="T">Tipo do model</typeparam>
        /// <param name="controller">Nome do controller da web api</param>
        /// <param name="param">Parâmetro(s) de busca do model</param>
        /// <returns></returns>
        public async Task<List<T>> GetListAsync<T>(string controller, string param = "") where T : class
        {
            string JsonResult = "";
            try
            {
                var client = new HttpClient();

                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(MediaType));

                string url = string.Format(ApiUrl + "/{0}?type=json&{1}", controller, param);
                var response = await client.GetAsync(url);

                JsonResult = response.Content.ReadAsStringAsync().Result;
                var rootobject = JsonConvert.DeserializeObject<List<T>>(JsonResult);

                return rootobject;
            }
            catch (Exception e)
            {
                try
                {
                    if (!String.IsNullOrEmpty(JsonResult))
                    {
                        APIErrorModel errorModel = JsonConvert.DeserializeObject<APIErrorModel>(JsonResult);
                        throw new Exception(errorModel.Message, new Exception(errorModel.ExceptionMessage));
                    }
                    else throw e;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        /// <summary>
        /// Atualiza o model
        /// </summary>
        /// <typeparam name="T">Tipo do model</typeparam>
        /// <param name="controller">Nome do controller da web api</param>
        /// <param name="param">Parâmetros de busca do model (chaves)</param>
        /// <param name="item">Model a ser atualizado</param>
        /// <returns></returns>
        public async Task<string> PutAsync<T>(string controller, string param, T item) where T : class
        {
            var client = new HttpClient();

            var json = JsonConvert.SerializeObject(item);
            var content = new StringContent(json, Encoding.UTF8, MediaType);
            HttpResponseMessage response = await client.PutAsync(string.Format(ApiUrl + "/{0}?{1}", controller, param), content);
            if (response.IsSuccessStatusCode)
            {
                return String.Empty;
            }
            else
            {
                string retorno = await response.Content.ReadAsStringAsync();
                try
                {
                    var rootobject = JsonConvert.DeserializeObject<APIJsonModel>(retorno);
                    return rootobject.ExceptionMessage;
                }
                catch (Exception ex)
                {
                    return retorno;
                }
            }
        }

        /// <summary>
        /// Inclui model
        /// </summary>
        /// <typeparam name="T">Tipo do model</typeparam>
        /// <param name="controller">Nome do controller da web api</param>
        /// <param name="item">Model a ser incluido</param>
        /// <returns></returns>
        public async Task<string> PostAsync<T>(string controller, T item) where T : class
        {
            var client = new HttpClient();

            var json = JsonConvert.SerializeObject(item);
            var content = new StringContent(json, Encoding.UTF8, MediaType);
            if (!String.IsNullOrEmpty(controller))
            {
                ApiUrl = string.Format(ApiUrl + "/{0}", controller);
            }

            HttpResponseMessage response = await client.PostAsync(ApiUrl, content);
            if (response.IsSuccessStatusCode)
            {
                return "";
            }
            else
            {
                var rootobject = JsonConvert.DeserializeObject<APIJsonModel>(await response.Content.ReadAsStringAsync());
                return rootobject.ExceptionMessage;
            }
        }

        public async Task<string> DeleteAsync(string controller, string param)
        {
            string url = string.Format(ApiUrl + "/{0}/{1}", controller, param);

            var client = new HttpClient();

            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(MediaType));

            var response = await client.DeleteAsync(url);

            if (response.IsSuccessStatusCode)
            {
                return "";
            }
            else
            {
                var rootobject = JsonConvert.DeserializeObject<APIJsonModel>(await response.Content.ReadAsStringAsync());
                return rootobject.ExceptionMessage;
            }
        }
    }
}