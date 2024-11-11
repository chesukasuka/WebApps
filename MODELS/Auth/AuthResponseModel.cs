namespace WebApps.MODELS.Auth
{
    public class AuthResponseModel
    {
        public string? token {  get; set; }
        public string? user_email {  get; set; }
        public string? user_nicename { get; set; }
        public string? user_display_name { get; set; }
    }
}
