using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SBO.Hub.Models.ServiceLayer;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Security;
using System.Text;
using System.Threading.Tasks;

namespace SBO.Hub.SBOHelpers
{
    public class ServiceLayer
    {
        private string ApiUrl;
        private static string SessionId;
        private static DateTime SessionTimeOut;
        private bool AutoLogin;
        private string Database;
        public string LastError;

        #region Constructors
        public ServiceLayer(bool automaticLogin = true, string database = "")
        {
            AutoLogin = automaticLogin;
            Database = database;

            ApiUrl = System.Configuration.ConfigurationManager.AppSettings["ServiceLayerURL"];
            if (String.IsNullOrEmpty(ApiUrl))
            {
                throw new Exception("URL da API deve ser informada no arquivo de configuração como no AppSettings como 'ServiceLayerURL' ou ser passada por parâmetro no construtor");
            }

            if (!ApiUrl.EndsWith("/"))
            {
                ApiUrl = ApiUrl + "/";
            }

            ServicePointManager.ServerCertificateValidationCallback = new
            RemoteCertificateValidationCallback
            (
               delegate { return true; }
            );
        }

        public ServiceLayer(string apiURL, bool automaticLogin = true)
        {
            ApiUrl = apiURL;

            if (!ApiUrl.EndsWith("/"))
            {
                ApiUrl = ApiUrl + "/";
            }

            ServicePointManager.ServerCertificateValidationCallback = new
            RemoteCertificateValidationCallback
            (
               delegate { return true; }
            );
        }
        #endregion

        #region SynchronousMethods
        public string Login(LoginModel loginModel = null)
        {
            return Task.Run(async () =>
            {
                return await LoginAsync(loginModel);
            }).Result;
        }
        public string Logout()
        {
            return Task.Run(async () =>
            {
                return await LogoutAsync();
            }).Result;
        }

        public T GetByID<T>(string entity, int id) where T : class
        {
            return Task.Run(async () =>
            {
                return await GetByIDAsync<T>(entity, id);
            }).Result;
        }

        public T GetByID<T>(string entity, string id) where T : class
        {
            return Task.Run(async () =>
            {
                return await GetByIDAsync<T>(entity, id);
            }).Result;
        }

        public T Get<T>(string entity, string filter) where T : class
        {
            return Task.Run(async () =>
            {
                return await GetAsync<T>(entity, filter);
            }).Result;
        }

        public List<T> GetList<T>(string entity, string filter) where T : class
        {
            return Task.Run(async () =>
            {
                return await GetListAsync<T>(entity, filter);
            }).Result;
        }

        public string Patch<T>(string entity, int id, T item) where T : class
        {
            return Task.Run(async () =>
            {
                return await PatchAsync<T>(entity, id, item);
            }).Result;
        }

        public string Patch<T>(string entity, string id, T item) where T : class
        {
            return Task.Run(async () =>
            {
                return await PatchAsync<T>(entity, id, item);
            }).Result;
        }

        public string Post<T>(string entity, T item) where T : class
        {
            return Task.Run(async () =>
            {
                return await PostAsync<T>(entity, item);
            }).Result;
        }

        public T PostAndGetAdded<T>(string entity, string newKeyFieldName, T item, bool retryInError = true) where T : class
        {
            return Task.Run(async () =>
            {
                return await PostAndGetAddedAsync<T>(entity, newKeyFieldName, item, retryInError);
            }).Result;
        }

        public string Delete(string entity, int id)
        {
            return Task.Run(async () =>
            {
                return await DeleteAsync(entity, id);
            }).Result;
        }

        public string Delete(string entity, string id)
        {
            return Task.Run(async () =>
            {
                return await DeleteAsync(entity, id);
            }).Result;
        }
        #endregion

        #region AsynchronousMethods
        public async Task<string> LoginAsync(LoginModel loginModel = null)
        {
            if (loginModel == null)
            {
                loginModel = new LoginModel();
                if (String.IsNullOrEmpty(Database))
                {
                    loginModel.CompanyDB = System.Configuration.ConfigurationManager.AppSettings["Database"];
                }
                else
                {
                    loginModel.CompanyDB = Database;
                }
                loginModel.UserName = System.Configuration.ConfigurationManager.AppSettings["B1User"];
                loginModel.Password = System.Configuration.ConfigurationManager.AppSettings["B1Password"];
            }
            try
            {
                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.ExpectContinue = false;

                var json = JsonConvert.SerializeObject(loginModel);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                HttpResponseMessage response = await client.PostAsync(ApiUrl + "Login", content);

                if (response.IsSuccessStatusCode)
                {
                    var rootobject = JsonConvert.DeserializeObject<SessionModel>(await response.Content.ReadAsStringAsync(), new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
                    SessionId = rootobject.SessionId;
                    SessionTimeOut = DateTime.Now.AddMinutes(rootobject.SessionTimeOut).AddMinutes(-1);
                    return "";
                }
                else
                {
                    return HandleError(await response.Content.ReadAsStringAsync());
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public async Task<string> LogoutAsync()
        {
            Uri uri = new Uri(string.Format(ApiUrl + "Logout"));
            HttpClientHandler handler = new HttpClientHandler();
            handler.CookieContainer = new CookieContainer();
            handler.CookieContainer.Add(uri, new Cookie("B1SESSION", SessionId));
            HttpClient client = new HttpClient(handler);

            client.DefaultRequestHeaders.ExpectContinue = false;

            var content = new StringContent("", Encoding.UTF8, "application/json");
            HttpResponseMessage response = await client.PostAsync(uri, content);
            if (response.IsSuccessStatusCode)
            {
                SessionId = "";
                return "";
            }
            else
            {
                return HandleError(await response.Content.ReadAsStringAsync());
            }
        }

        public async Task<T> GetByIDAsync<T>(string entity, int id) where T : class
        {
            string url = $"{ApiUrl}{entity}({id})";
            return await this.GetByIDAsync<T>(url);
        }

        public async Task<T> GetByIDAsync<T>(string entity, string id) where T : class
        {
            string url = $"{ApiUrl}{entity}('{id}')";
            return await this.GetByIDAsync<T>(url);
        }

        public async Task<T> GetByIDAsync<T>(string url) where T : class
        {
            string JsonResult = String.Empty;
            try
            {
                if (AutoLogin && DateTime.Now > SessionTimeOut)
                {
                    await this.LoginAsync();
                }

                HttpClientHandler handler = new HttpClientHandler();
                handler.CookieContainer = new CookieContainer();
                handler.CookieContainer.Add(new Uri(url), new Cookie("B1SESSION", SessionId));

                HttpClient client = new HttpClient(handler);
                client.DefaultRequestHeaders.ExpectContinue = false;

                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                HttpResponseMessage response = await client.GetAsync(url);

                JsonResult = response.Content.ReadAsStringAsync().Result;
                var rootobject = JsonConvert.DeserializeObject<T>(JsonResult, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });

                return rootobject;
            }
            catch (Exception e)
            {
                if (!String.IsNullOrEmpty(JsonResult))
                {
                    string error = HandleError(JsonResult);
                    if (!String.IsNullOrEmpty(error))
                    {
                        throw new Exception(error);

                    }
                }
                throw e;
            }
        }

        public async Task<T> GetAsync<T>(string entity, string filter) where T : class
        {
            string JsonResult = String.Empty;
            try
            {
                if (AutoLogin && DateTime.Now > SessionTimeOut)
                {
                    await this.LoginAsync();
                }

                string url = $"{ApiUrl}{entity}?$filter={filter}";
                HttpClientHandler handler = new HttpClientHandler();
                handler.CookieContainer = new CookieContainer();
                handler.CookieContainer.Add(new Uri(url), new Cookie("B1SESSION", SessionId));

                HttpClient client = new HttpClient(handler);
                client.DefaultRequestHeaders.ExpectContinue = false;

                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                HttpResponseMessage response = await client.GetAsync(url);

                JsonResult = response.Content.ReadAsStringAsync().Result;
                var rootobject = JsonConvert.DeserializeObject<T>(JsonResult, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });

                return rootobject;
            }
            catch (Exception e)
            {
                if (!String.IsNullOrEmpty(JsonResult))
                {
                    string error = HandleError(JsonResult);
                    if (!String.IsNullOrEmpty(error))
                    {
                        throw new Exception(error);

                    }
                }
                throw e;
            }
        }

        public async Task<List<T>> GetListAsync<T>(string entity, string filter = "") where T : class
        {
            string JsonResult = String.Empty;
            try
            {
                if (AutoLogin && DateTime.Now > SessionTimeOut)
                {
                    await this.LoginAsync();
                }

                string url = $"{ApiUrl}{entity}?$filter={filter}";
                HttpClientHandler handler = new HttpClientHandler();
                handler.CookieContainer = new CookieContainer();
                handler.CookieContainer.Add(new Uri(url), new Cookie("B1SESSION", SessionId));

                HttpClient client = new HttpClient(handler);
                client.DefaultRequestHeaders.ExpectContinue = false;

                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                var response = await client.GetAsync(url);

                JsonResult = response.Content.ReadAsStringAsync().Result;
                if (entity.StartsWith("U_")) // Se é tabela de usuário, é colocado dentro do campo "value"
                {
                    JObject obj = JObject.Parse(JsonResult);
                    var modelObject = JsonConvert.DeserializeObject<List<T>>(obj["value"].ToString(), new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
                    return modelObject;
                }
                else
                {
                    var modelObject = JsonConvert.DeserializeObject<List<T>>(JsonResult, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
                    return modelObject;
                }
            }
            catch (Exception e)
            {
                if (!String.IsNullOrEmpty(JsonResult))
                {
                    string error = HandleError(JsonResult);
                    if (!String.IsNullOrEmpty(error))
                    {
                        throw new Exception(error);

                    }
                }
                throw e;
            }
        }

        public async Task<string> PatchAsync<T>(string entity, int param, T item) where T : class
        {
            Uri uri = new Uri($"{ApiUrl}{entity}({param})");
            return await PatchAsync<T>(uri, item);
        }

        public async Task<string> PatchAsync<T>(string entity, string param, T item) where T : class
        {
            Uri uri = new Uri($"{ApiUrl}{entity}('{param}')");
            return await PatchAsync<T>(uri, item);
        }

        public async Task<string> PatchAsync<T>(Uri uri, T item) where T : class
        {
            if (AutoLogin && DateTime.Now > SessionTimeOut)
            {
                await this.LoginAsync();
            }

            HttpClientHandler handler = new HttpClientHandler() { UseCookies = false };
            HttpClient client = new HttpClient(handler);
            client.BaseAddress = uri;
            client.DefaultRequestHeaders.ExpectContinue = false;

            var json = JsonConvert.SerializeObject(item, Formatting.None, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var method = new HttpMethod("PATCH");
            var request = new HttpRequestMessage(method, uri) { Content = content };
            request.Headers.Add("Cookie", "B1SESSION=" + SessionId);

            HttpResponseMessage response = await client.SendAsync(request);
            if (response.IsSuccessStatusCode)
            {
                return String.Empty;
            }
            else
            {
                return HandleError(await response.Content.ReadAsStringAsync());
            }
        }

        public async Task<string> PostAsync<T>(string entity, T item) where T : class
        {
            try
            {
                if (AutoLogin && DateTime.Now > SessionTimeOut)
                {
                    await this.LoginAsync();
                }

                Uri uri = new Uri($"{ApiUrl}{entity}");

                HttpClientHandler handler = new HttpClientHandler();
                handler.CookieContainer = new CookieContainer();
                handler.CookieContainer.Add(new Cookie("B1SESSION", SessionId) { Domain = uri.Host });
                HttpClient client = new HttpClient(handler);

                var json = JsonConvert.SerializeObject(item, Formatting.None, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });

                var content = new StringContent(json, Encoding.UTF8, "application/json");
                HttpResponseMessage response = await client.PostAsync(uri, content);

                if (response.IsSuccessStatusCode)
                {
                    json = await response.Content.ReadAsStringAsync();
                    item = JsonConvert.DeserializeObject<T>(json);
                    return "";
                }
                else
                {
                    return HandleError(await response.Content.ReadAsStringAsync());
                }
            }
            catch (Exception ex)
            {
                LastError = ex.Message;
                return ex.Message;
            }
        }


        public async Task<T> PostAndGetAddedAsync<T>(string entity, string newKeyFieldName, T item, bool retryInError = true) where T : class
        {
            try
            {
                if (AutoLogin && DateTime.Now > SessionTimeOut)
                {
                    await this.LoginAsync();
                }

                Uri uri = new Uri($"{ApiUrl}{entity}");

                HttpClientHandler handler = new HttpClientHandler();
                handler.CookieContainer = new CookieContainer();
                handler.CookieContainer.Add(new Cookie("B1SESSION", SessionId) { Domain = uri.Host });
                HttpClient client = new HttpClient(handler);

                var json = JsonConvert.SerializeObject(item, Formatting.None, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });

                var content = new StringContent(json, Encoding.UTF8, "application/json");
                HttpResponseMessage response = await client.PostAsync(uri, content);
                json = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    if (String.IsNullOrEmpty(json))
                    {
                        if (retryInError)
                        {
                            await this.LoginAsync();
                            return await PostAndGetAddedAsync<T>(entity, newKeyFieldName, item, false);
                        }
                    }

                    item = JsonConvert.DeserializeObject<T>(json);
                    return item;
                }
                else
                {
                    if (String.IsNullOrEmpty(json))
                    {
                        if (retryInError)
                        {
                            await this.LoginAsync();
                            return await PostAndGetAddedAsync<T>(entity, newKeyFieldName, item, false);
                        }
                    }

                    throw new Exception(HandleError(json));
                }
            }
            catch (Exception ex)
            {
                LastError = ex.Message;
                throw ex;
            }
        }

        public async Task<string> DeleteAsync(string entity, string id)
        {
            string url = $"{ApiUrl}{entity}('{id}')";

            return await DeleteAsync(url);
        }

        public async Task<string> DeleteAsync(string entity, int id)
        {
            string url = $"{ApiUrl}{entity}({id})";
            return await DeleteAsync(url);
        }

        public async Task<string> DeleteAsync(string url)
        {
            if (AutoLogin && DateTime.Now > SessionTimeOut)
            {
                await this.LoginAsync();
            }

            HttpClientHandler handler = new HttpClientHandler();
            handler.CookieContainer = new CookieContainer();
            handler.CookieContainer.Add(new Cookie("B1SESSION", SessionId));
            HttpClient client = new HttpClient(handler);

            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            var response = await client.DeleteAsync(url);

            if (response.IsSuccessStatusCode)
            {
                return "";
            }
            else
            {
                return HandleError(await response.Content.ReadAsStringAsync());
            }
        }
        #endregion

        private string HandleError(string response)
        {
            LastError = response;
            try
            {
                ServiceLayerBaseModel serviceLayerBaseModel = JsonConvert.DeserializeObject<ServiceLayerBaseModel>(response);
                if (!String.IsNullOrEmpty(serviceLayerBaseModel.ExceptionMessage))
                {
                    LastError = serviceLayerBaseModel.ExceptionMessage;
                    return serviceLayerBaseModel.ExceptionMessage;
                }
                SBOErrorModel sboErrorModel = JsonConvert.DeserializeObject<SBOErrorModel>(response);
                if (sboErrorModel.error != null)
                {
                    LastError = sboErrorModel.error.code + " - " + sboErrorModel.error.message.value;
                    return sboErrorModel.error.code + " - " + sboErrorModel.error.message.value;
                }
            }
            catch (Exception ex)
            {
                return response;
            }
            return response;
        }
    }
}
