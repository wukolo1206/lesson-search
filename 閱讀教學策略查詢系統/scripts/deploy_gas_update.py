import subprocess
import json

def update_gas_deployment():
    print("Fetching deployments...")
    try:
        # Get list of deployments
        result = subprocess.run(
            ['clasp', 'deployments'], 
            capture_output=True, 
            text=True, 
            check=True
        )
        output_lines = result.stdout.strip().split('\n')
        
        # Look for the web app deployment id
        deployment_id = None
        for line in output_lines:
            if '- @' in line and 'web app' in line.lower():
                # example line: - AKfycbz8T4Sq... @3 - web app version
                parts = line.split('-')[1].strip().split(' ')
                deployment_id = parts[0]
                break
        
        if not deployment_id:
            print("找不到現有的 Web App 部署 ID，將會建立新的部署。")
            subprocess.run(['clasp', 'deploy'], check=True)
            print("新部署完成！")
            return
            
        print(f"找到原本的部署 ID: {deployment_id}")
        print("正在更新部署...")
        subprocess.run(['clasp', 'deploy', '-i', deployment_id], check=True)
        print("部署更新完成！您的網址維持不變。")
        
    except subprocess.CalledProcessError as e:
        print(f"Error executing clasp command: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    update_gas_deployment()
