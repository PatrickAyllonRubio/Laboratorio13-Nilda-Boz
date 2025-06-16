using Laboratorio13___Nilda_Boza.Infrastructure.Services;
using Microsoft.Extensions.DependencyInjection;

namespace Laboratorio13___Nilda_Boza.Infrastructure.Configuration;

public static class InfrastructureServicesExtensions
{
    public static IServiceCollection AddInfrastructureServices(this IServiceCollection services) 
    { 
    
        
        //services.AddTransient<IUnitOfWork, UnitOfWork>();
        //services.AddScoped(typeof(IGenericRepository<,>), typeof(GenericRepository<,>));
        services.AddScoped<ExcelExportService>();
        
        
        return services;
    }
}