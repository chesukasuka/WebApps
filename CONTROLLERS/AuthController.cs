using Microsoft.AspNetCore.Identity.Data;
using Microsoft.AspNetCore.Mvc;
using Microsoft.IdentityModel.Tokens;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;
using WebApps.MODELS.Auth;

namespace WebApps.CONTROLLERS
{
    public class AuthController : Controller
    {

        [HttpPost]
        public async Task<IActionResult> Login([FromForm] AuthRequestModel request)
        {
            using var httpClient = new HttpClient();
            var response = await httpClient.PostAsync($"https://centraldataaccess.co.id/wp-json/jwt-auth/v1/token?username={request.username}&password={request.password}", null);

            if (response.IsSuccessStatusCode)
            {
                var result = await response.Content.ReadFromJsonAsync<AuthResponseModel>();
                HttpContext.Session.SetString("Token", result.token);
                HttpContext.Session.SetString("Email", result.user_email);
                HttpContext.Session.SetString("Company", result.user_nicename);
                HttpContext.Session.SetString("Name", result.user_display_name);
                return RedirectToAction("Benchmarking", "Service");
            }
            var errresult = await response.Content.ReadFromJsonAsync<AuthErrResponseModel>();
            return RedirectToAction("Benchmarking", "Service", new { errMsg = errresult.message });
        }

        [HttpGet]
        public IActionResult Logout()
        {
            var hitungCount = HttpContext.Session.GetString("HitungCount");
            HttpContext.Session.Clear();
            if(hitungCount != null)
            {
                HttpContext.Session.SetString("HitungCount", hitungCount);
            }
            return RedirectToAction("Benchmarking", "Service");
        }
    }
}
