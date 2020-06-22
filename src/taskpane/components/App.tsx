import * as React from "react";
import Progress from "./Progress";
import { Button, Divider, Row, Descriptions, List } from "antd";
import "antd/dist/antd.css";
import { loginService } from "../../services/services";
import useObservable from "../../hooks/useObservable";
import { SharePointController, GraphController, AzureController } from "../../controllers";
import { Spin, Space } from "antd";
import { Alert } from "antd";
import Search from "antd/lib/input/Search";
import { User } from "@microsoft/microsoft-graph-types";

interface IAppProps {
  title: string;
  isOfficeInitialized: boolean;
}

const App: React.FC<IAppProps> = props => {
  const tokens = useObservable(loginService.tokens);
  const error = useObservable(loginService.errorMessage);
  const [userInfo, setUserInfo] = React.useState<User>(null);
  const [azureResponse, setAzureResponse] = React.useState("");
  const [searchResult, setSearchResult] = React.useState([]);
  const [loading, setLoading] = React.useState(false);

  const getUserDetails = async () => {
    const controller = new GraphController(tokens.graphToken);
    const user = await controller.getUserInformation();
    setUserInfo(user);
  };

  const searchDocuments = async value => {
    const spController = new SharePointController(tokens.sharePointToken);
    const result = await spController.searchDocuments(value);
    setSearchResult(result);
  };

  const callAzureFunction = async value => {
    const azureController = new AzureController(tokens.azureToken);
    const response = await azureController.callAzureFunction(value);
    setAzureResponse(response);
  };

  const logOut = () => {
    setLoading(false);
    loginService.logOut();
  };

  const logIn = () => {
    setLoading(true);
    loginService.getAccessToken();
  };

  const onClose = (e: React.MouseEvent<HTMLButtonElement, MouseEvent>) => {
    console.log(e, "Error message closed!");
  };

  if (!props.isOfficeInitialized) {
    return (
      <Progress
        title={props.title}
        logo="assets/logo-filled.png"
        message="Please sideload your addin to see app body."
      />
    );
  }
  if (!tokens) {
    return (
      <div className="ms-welcome__header ms-u-fadeIn500">
        {error && <Alert message="Error" description={error} type="error" closable onClose={onClose} />}
        {!loading && (
          <section className="ms-welcome__header ms-u-fadeIn500">
            <h1 className="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary">
              Welcome!
            </h1>
            <h3>Please sign in to see the content.</h3>
            <Button type="dashed" danger onClick={logIn}>
              Log in
            </Button>
            
          </section>
        )}
        {loading && (
          <Space size="middle">
            <Spin size="large" tip="Loading..." />
          </Space>
        )}
      </div>
    );
  }
  return (
    <div className="ms-welcome__main">
      <Row className="centerBlock">
        <Button type="dashed" danger onClick={logOut}>
          Log Out
        </Button>
      </Row>
      <Divider>Graph API</Divider>
      <Space direction="vertical">
        {!userInfo && (
          <Button type="primary" onClick={getUserDetails}>
            Get User Profile
          </Button>
        )}

        {userInfo && (
          <Descriptions>
            <Descriptions.Item label="First Name">{userInfo?.givenName}</Descriptions.Item>
            <Descriptions.Item label="Last Name">{userInfo?.surname}</Descriptions.Item>
            <Descriptions.Item label="Job Title">{userInfo?.jobTitle}</Descriptions.Item>
            <Descriptions.Item label="Department">{userInfo?.department}</Descriptions.Item>
            <Descriptions.Item label="Mobile">{userInfo?.mobilePhone}</Descriptions.Item>
            <Descriptions.Item label="Phone">{userInfo?.businessPhones[0]}</Descriptions.Item>
            <Descriptions.Item label="City">{userInfo?.city}</Descriptions.Item>
          </Descriptions>
        )}
      </Space>

      <Divider>SharePoint API</Divider>
      <Search placeholder="Search for documents in SharePoint" onSearch={value => searchDocuments(value)} enterButton />
      {searchResult.length > 0 && (
        <Row>
          <List
            itemLayout="horizontal"
            dataSource={searchResult}
            renderItem={item => (
              <List.Item>
                <a href={item.Path}>{item.Title}</a>
              </List.Item>
            )}
          />
        </Row>
      )}
      <Divider>Azure Function</Divider>
      <Search
        placeholder="Enter your name"
        enterButton="Get"
        size="middle"
        onSearch={value => callAzureFunction(value)}
      />
      <span className="azure-result-span">{azureResponse}</span>
    </div>
  );
};

export default App;
